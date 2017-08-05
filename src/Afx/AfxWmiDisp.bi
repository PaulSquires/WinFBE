' ########################################################################################
' Microsoft Windows
' File: AfxWmiDisp.bi
' Library name: WbemScripting
' Documentation string: Microsoft WMI Scripting V1.2 Library
' GUID: {565783C6-CB41-11D1-8B02-00600806D9B6}
' Version: 1.2, Locale ID = 0
' Path: C:\Windows\system32\wbem\wbemdisp.TLB
' Attributes: 8 [&h00000008]  [HasDiskImage]
' ########################################################################################

#pragma once
#include once "win/control.bi"
#include once "win/wbemcli.bi"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' // ProgIDs (Program Identifiers)
CONST AFX_PROGID_WbemScripting_SWbemDateTime_1 = "WbemScripting.SWbemDateTime.1"
CONST AFX_PROGID_WbemScripting_SWbemLastError_1 = "WbemScripting.SWbemLastError.1"
CONST AFX_PROGID_WbemScripting_SWbemLocator_1 = "WbemScripting.SWbemLocator.1"
CONST AFX_PROGID_WbemScripting_SWbemNamedValueSet_1 = "WbemScripting.SWbemNamedValueSet.1"
CONST AFX_PROGID_WbemScripting_SWbemObjectPath_1 = "WbemScripting.SWbemObjectPath.1"
CONST AFX_PROGID_WbemScripting_SWbemRefresher_1 = "WbemScripting.SWbemRefresher.1"
CONST AFX_PROGID_WbemScripting_SWbemSink_1 = "WbemScripting.SWbemSink.1"

' // Version independent ProgIDs
CONST AFX_PROGID_WbemScripting_SWbemDateTime = "WbemScripting.SWbemDateTime"
CONST AFX_PROGID_WbemScripting_SWbemLastError = "WbemScripting.SWbemLastError"
CONST AFX_PROGID_WbemScripting_SWbemLocator = "WbemScripting.SWbemLocator"
CONST AFX_PROGID_WbemScripting_SWbemNamedValueSet = "WbemScripting.SWbemNamedValueSet"
CONST AFX_PROGID_WbemScripting_SWbemObjectPath = "WbemScripting.SWbemObjectPath"
CONST AFX_PROGID_WbemScripting_SWbemRefresher = "WbemScripting.SWbemRefresher"
CONST AFX_PROGID_WbemScripting_SWbemSink = "WbemScripting.SWbemSink"

' // ClsIDs (Class identifiers)
CONST AFX_CLSID_SWbemDateTime = "{47DFBE54-CF76-11D3-B38F-00105A1F473A}"
CONST AFX_CLSID_SWbemEventSource = "{04B83D58-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemLastError = "{C2FEEEAC-CFCD-11D1-8B05-00600806D9B6}"
CONST AFX_CLSID_SWbemLocator = "{76A64158-CB41-11D1-8B02-00600806D9B6}"
CONST AFX_CLSID_SWbemMethod = "{04B83D5B-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemMethodSet = "{04B83D5A-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemNamedValue = "{04B83D60-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemNamedValueSet = "{9AED384E-CE8B-11D1-8B05-00600806D9B6}"
CONST AFX_CLSID_SWbemObject = "{04B83D62-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemObjectEx = "{D6BDAFB2-9435-491F-BB87-6AA0F0BC31A2}"
CONST AFX_CLSID_SWbemObjectPath = "{5791BC26-CE9C-11D1-97BF-0000F81E849C}"
CONST AFX_CLSID_SWbemObjectSet = "{04B83D61-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemPrivilege = "{26EE67BC-5804-11D2-8B4A-00600806D9B6}"
CONST AFX_CLSID_SWbemPrivilegeSet = "{26EE67BE-5804-11D2-8B4A-00600806D9B6}"
CONST AFX_CLSID_SWbemProperty = "{04B83D5D-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemPropertySet = "{04B83D5C-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemQualifier = "{04B83D5F-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemQualifierSet = "{04B83D5E-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemRefreshableItem = "{8C6854BC-DE4B-11D3-B390-00105A1F473A}"
CONST AFX_CLSID_SWbemRefresher = "{D269BF5C-D9C1-11D3-B38F-00105A1F473A}"
CONST AFX_CLSID_SWbemSecurity = "{B54D66E9-2287-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemServices = "{04B83D63-21AE-11D2-8B33-00600806D9B6}"
CONST AFX_CLSID_SWbemServicesEx = "{62E522DC-8CF3-40A8-8B2E-37D595651E40}"
CONST AFX_CLSID_SWbemSink = "{75718C9A-F029-11D1-A1AC-00C04FB6C223}"

' // IIDs (Interface identifiers)
CONST AFX_IID_ISWbemDateTime = "{5E97458A-CF77-11D3-B38F-00105A1F473A}"
CONST AFX_IID_ISWbemEventSource = "{27D54D92-0EBE-11D2-8B22-00600806D9B6}"
CONST AFX_IID_ISWbemLastError = "{D962DB84-D4BB-11D1-8B09-00600806D9B6}"
CONST AFX_IID_ISWbemLocator = "{76A6415B-CB41-11D1-8B02-00600806D9B6}"
CONST AFX_IID_ISWbemMethod = "{422E8E90-D955-11D1-8B09-00600806D9B6}"
CONST AFX_IID_ISWbemMethodSet = "{C93BA292-D955-11D1-8B09-00600806D9B6}"
CONST AFX_IID_ISWbemNamedValue = "{76A64164-CB41-11D1-8B02-00600806D9B6}"
CONST AFX_IID_ISWbemNamedValueSet = "{CF2376EA-CE8C-11D1-8B05-00600806D9B6}"
CONST AFX_IID_ISWbemObject = "{76A6415A-CB41-11D1-8B02-00600806D9B6}"
CONST AFX_IID_ISWbemObjectEx = "{269AD56A-8A67-4129-BC8C-0506DCFE9880}"
CONST AFX_IID_ISWbemObjectPath = "{5791BC27-CE9C-11D1-97BF-0000F81E849C}"
CONST AFX_IID_ISWbemObjectSet = "{76A6415F-CB41-11D1-8B02-00600806D9B6}"
CONST AFX_IID_ISWbemPrivilege = "{26EE67BD-5804-11D2-8B4A-00600806D9B6}"
CONST AFX_IID_ISWbemPrivilegeSet = "{26EE67BF-5804-11D2-8B4A-00600806D9B6}"
CONST AFX_IID_ISWbemProperty = "{1A388F98-D4BA-11D1-8B09-00600806D9B6}"
CONST AFX_IID_ISWbemPropertySet = "{DEA0A7B2-D4BA-11D1-8B09-00600806D9B6}"
CONST AFX_IID_ISWbemQualifier = "{79B05932-D3B7-11D1-8B06-00600806D9B6}"
CONST AFX_IID_ISWbemQualifierSet = "{9B16ED16-D3DF-11D1-8B08-00600806D9B6}"
CONST AFX_IID_ISWbemRefreshableItem = "{5AD4BF92-DAAB-11D3-B38F-00105A1F473A}"
CONST AFX_IID_ISWbemRefresher = "{14D8250E-D9C2-11D3-B38F-00105A1F473A}"
CONST AFX_IID_ISWbemSecurity = "{B54D66E6-2287-11D2-8B33-00600806D9B6}"
CONST AFX_IID_ISWbemServices = "{76A6415C-CB41-11D1-8B02-00600806D9B6}"
CONST AFX_IID_ISWbemServicesEx = "{D2F68443-85DC-427E-91D8-366554CC754C}"
CONST AFX_IID_ISWbemSink = "{75718C9F-F029-11D1-A1AC-00C04FB6C223}"
CONST AFX_IID_ISWbemSinkEvents = "{75718CA0-F029-11D1-A1AC-00C04FB6C223}"

' // Enumerations

ENUM WbemAuthenticationLevelEnum
   ' // IID: {B54D66E7-2287-11D2-8B33-00600806D9B6}
   ' // Documentation string: Defines the security authentication level
   ' // Number of constants: 7
   wbemAuthenticationLevelDefault = 0   ' (&h00000000)
   wbemAuthenticationLevelNone = 1   ' (&h00000001)
   wbemAuthenticationLevelConnect = 2   ' (&h00000002)
   wbemAuthenticationLevelCall = 3   ' (&h00000003)
   wbemAuthenticationLevelPkt = 4   ' (&h00000004)
   wbemAuthenticationLevelPktIntegrity = 5   ' (&h00000005)
   wbemAuthenticationLevelPktPrivacy = 6   ' (&h00000006)
END ENUM

ENUM WbemChangeFlagEnum
   ' // IID: {4A249B72-FC9A-11D1-8B1E-00600806D9B6}
   ' // Documentation string: Defines semantics of putting a Class or Instance
   ' // Number of constants: 8
   wbemChangeFlagCreateOrUpdate = 0   ' (&h00000000)
   wbemChangeFlagUpdateOnly = 1   ' (&h00000001)
   wbemChangeFlagCreateOnly = 2   ' (&h00000002)
   wbemChangeFlagUpdateCompatible = 0   ' (&h00000000)
   wbemChangeFlagUpdateSafeMode = 32   ' (&h00000020)
   wbemChangeFlagUpdateForceMode = 64   ' (&h00000040)
   wbemChangeFlagStrongValidation = 128   ' (&h00000080)
   wbemChangeFlagAdvisory = 65536   ' (&h00010000)
END ENUM

ENUM WbemCimtypeEnum
   ' // IID: {4A249B7B-FC9A-11D1-8B1E-00600806D9B6}
   ' // Documentation string: Defines the valid CIM Types of a Property value
   ' // Number of constants: 16
   wbemCimtypeSint8 = 16   ' (&h00000010)
   wbemCimtypeUint8 = 17   ' (&h00000011)
   wbemCimtypeSint16 = 2   ' (&h00000002)
   wbemCimtypeUint16 = 18   ' (&h00000012)
   wbemCimtypeSint32 = 3   ' (&h00000003)
   wbemCimtypeUint32 = 19   ' (&h00000013)
   wbemCimtypeSint64 = 20   ' (&h00000014)
   wbemCimtypeUint64 = 21   ' (&h00000015)
   wbemCimtypeReal32 = 4   ' (&h00000004)
   wbemCimtypeReal64 = 5   ' (&h00000005)
   wbemCimtypeBoolean = 11   ' (&h0000000B)
   wbemCimtypeString = 8   ' (&h00000008)
   wbemCimtypeDatetime = 101   ' (&h00000065)
   wbemCimtypeReference = 102   ' (&h00000066)
   wbemCimtypeChar16 = 103   ' (&h00000067)
   wbemCimtypeObject = 13   ' (&h0000000D)
END ENUM

ENUM WbemComparisonFlagEnum
   ' // IID: {4A249B79-FC9A-11D1-8B1E-00600806D9B6}
   ' // Documentation string: Defines settings for object comparison
   ' // Number of constants: 7
   wbemComparisonFlagIncludeAll = 0   ' (&h00000000)
   wbemComparisonFlagIgnoreQualifiers = 1   ' (&h00000001)
   wbemComparisonFlagIgnoreObjectSource = 2   ' (&h00000002)
   wbemComparisonFlagIgnoreDefaultValues = 4   ' (&h00000004)
   wbemComparisonFlagIgnoreClass = 8   ' (&h00000008)
   wbemComparisonFlagIgnoreCase = 16   ' (&h00000010)
   wbemComparisonFlagIgnoreFlavor = 32   ' (&h00000020)
END ENUM

ENUM WbemConnectOptionsEnum
   ' // Documentation string: Used to define connection behavior
   ' // Number of constants: 1
   wbemConnectFlagUseMaxWait = 128   ' (&h00000080)
END ENUM

ENUM WbemErrorEnum
   ' // IID: {4A249B7C-FC9A-11D1-8B1E-00600806D9B6}
   ' // Documentation string: Defines the errors that may be returned by the WBEM Scripting library
   ' // Number of constants: 128
   wbemNoErr = 0   ' (&h00000000)
   wbemErrFailed = -2147217407   ' (&h80041001)
   wbemErrNotFound = -2147217406   ' (&h80041002)
   wbemErrAccessDenied = -2147217405   ' (&h80041003)
   wbemErrProviderFailure = -2147217404   ' (&h80041004)
   wbemErrTypeMismatch = -2147217403   ' (&h80041005)
   wbemErrOutOfMemory = -2147217402   ' (&h80041006)
   wbemErrInvalidContext = -2147217401   ' (&h80041007)
   wbemErrInvalidParameter = -2147217400   ' (&h80041008)
   wbemErrNotAvailable = -2147217399   ' (&h80041009)
   wbemErrCriticalError = -2147217398   ' (&h8004100A)
   wbemErrInvalidStream = -2147217397   ' (&h8004100B)
   wbemErrNotSupported = -2147217396   ' (&h8004100C)
   wbemErrInvalidSuperclass = -2147217395   ' (&h8004100D)
   wbemErrInvalidNamespace = -2147217394   ' (&h8004100E)
   wbemErrInvalidObject = -2147217393   ' (&h8004100F)
   wbemErrInvalidClass = -2147217392   ' (&h80041010)
   wbemErrProviderNotFound = -2147217391   ' (&h80041011)
   wbemErrInvalidProviderRegistration = -2147217390   ' (&h80041012)
   wbemErrProviderLoadFailure = -2147217389   ' (&h80041013)
   wbemErrInitializationFailure = -2147217388   ' (&h80041014)
   wbemErrTransportFailure = -2147217387   ' (&h80041015)
   wbemErrInvalidOperation = -2147217386   ' (&h80041016)
   wbemErrInvalidQuery = -2147217385   ' (&h80041017)
   wbemErrInvalidQueryType = -2147217384   ' (&h80041018)
   wbemErrAlreadyExists = -2147217383   ' (&h80041019)
   wbemErrOverrideNotAllowed = -2147217382   ' (&h8004101A)
   wbemErrPropagatedQualifier = -2147217381   ' (&h8004101B)
   wbemErrPropagatedProperty = -2147217380   ' (&h8004101C)
   wbemErrUnexpected = -2147217379   ' (&h8004101D)
   wbemErrIllegalOperation = -2147217378   ' (&h8004101E)
   wbemErrCannotBeKey = -2147217377   ' (&h8004101F)
   wbemErrIncompleteClass = -2147217376   ' (&h80041020)
   wbemErrInvalidSyntax = -2147217375   ' (&h80041021)
   wbemErrNondecoratedObject = -2147217374   ' (&h80041022)
   wbemErrReadOnly = -2147217373   ' (&h80041023)
   wbemErrProviderNotCapable = -2147217372   ' (&h80041024)
   wbemErrClassHasChildren = -2147217371   ' (&h80041025)
   wbemErrClassHasInstances = -2147217370   ' (&h80041026)
   wbemErrQueryNotImplemented = -2147217369   ' (&h80041027)
   wbemErrIllegalNull = -2147217368   ' (&h80041028)
   wbemErrInvalidQualifierType = -2147217367   ' (&h80041029)
   wbemErrInvalidPropertyType = -2147217366   ' (&h8004102A)
   wbemErrValueOutOfRange = -2147217365   ' (&h8004102B)
   wbemErrCannotBeSingleton = -2147217364   ' (&h8004102C)
   wbemErrInvalidCimType = -2147217363   ' (&h8004102D)
   wbemErrInvalidMethod = -2147217362   ' (&h8004102E)
   wbemErrInvalidMethodParameters = -2147217361   ' (&h8004102F)
   wbemErrSystemProperty = -2147217360   ' (&h80041030)
   wbemErrInvalidProperty = -2147217359   ' (&h80041031)
   wbemErrCallCancelled = -2147217358   ' (&h80041032)
   wbemErrShuttingDown = -2147217357   ' (&h80041033)
   wbemErrPropagatedMethod = -2147217356   ' (&h80041034)
   wbemErrUnsupportedParameter = -2147217355   ' (&h80041035)
   wbemErrMissingParameter = -2147217354   ' (&h80041036)
   wbemErrInvalidParameterId = -2147217353   ' (&h80041037)
   wbemErrNonConsecutiveParameterIds = -2147217352   ' (&h80041038)
   wbemErrParameterIdOnRetval = -2147217351   ' (&h80041039)
   wbemErrInvalidObjectPath = -2147217350   ' (&h8004103A)
   wbemErrOutOfDiskSpace = -2147217349   ' (&h8004103B)
   wbemErrBufferTooSmall = -2147217348   ' (&h8004103C)
   wbemErrUnsupportedPutExtension = -2147217347   ' (&h8004103D)
   wbemErrUnknownObjectType = -2147217346   ' (&h8004103E)
   wbemErrUnknownPacketType = -2147217345   ' (&h8004103F)
   wbemErrMarshalVersionMismatch = -2147217344   ' (&h80041040)
   wbemErrMarshalInvalidSignature = -2147217343   ' (&h80041041)
   wbemErrInvalidQualifier = -2147217342   ' (&h80041042)
   wbemErrInvalidDuplicateParameter = -2147217341   ' (&h80041043)
   wbemErrTooMuchData = -2147217340   ' (&h80041044)
   wbemErrServerTooBusy = -2147217339   ' (&h80041045)
   wbemErrInvalidFlavor = -2147217338   ' (&h80041046)
   wbemErrCircularReference = -2147217337   ' (&h80041047)
   wbemErrUnsupportedClassUpdate = -2147217336   ' (&h80041048)
   wbemErrCannotChangeKeyInheritance = -2147217335   ' (&h80041049)
   wbemErrCannotChangeIndexInheritance = -2147217328   ' (&h80041050)
   wbemErrTooManyProperties = -2147217327   ' (&h80041051)
   wbemErrUpdateTypeMismatch = -2147217326   ' (&h80041052)
   wbemErrUpdateOverrideNotAllowed = -2147217325   ' (&h80041053)
   wbemErrUpdatePropagatedMethod = -2147217324   ' (&h80041054)
   wbemErrMethodNotImplemented = -2147217323   ' (&h80041055)
   wbemErrMethodDisabled = -2147217322   ' (&h80041056)
   wbemErrRefresherBusy = -2147217321   ' (&h80041057)
   wbemErrUnparsableQuery = -2147217320   ' (&h80041058)
   wbemErrNotEventClass = -2147217319   ' (&h80041059)
   wbemErrMissingGroupWithin = -2147217318   ' (&h8004105A)
   wbemErrMissingAggregationList = -2147217317   ' (&h8004105B)
   wbemErrPropertyNotAnObject = -2147217316   ' (&h8004105C)
   wbemErrAggregatingByObject = -2147217315   ' (&h8004105D)
   wbemErrUninterpretableProviderQuery = -2147217313   ' (&h8004105F)
   wbemErrBackupRestoreWinmgmtRunning = -2147217312   ' (&h80041060)
   wbemErrQueueOverflow = -2147217311   ' (&h80041061)
   wbemErrPrivilegeNotHeld = -2147217310   ' (&h80041062)
   wbemErrInvalidOperator = -2147217309   ' (&h80041063)
   wbemErrLocalCredentials = -2147217308   ' (&h80041064)
   wbemErrCannotBeAbstract = -2147217307   ' (&h80041065)
   wbemErrAmendedObject = -2147217306   ' (&h80041066)
   wbemErrClientTooSlow = -2147217305   ' (&h80041067)
   wbemErrNullSecurityDescriptor = -2147217304   ' (&h80041068)
   wbemErrTimeout = -2147217303   ' (&h80041069)
   wbemErrInvalidAssociation = -2147217302   ' (&h8004106A)
   wbemErrAmbiguousOperation = -2147217301   ' (&h8004106B)
   wbemErrQuotaViolation = -2147217300   ' (&h8004106C)
   wbemErrTransactionConflict = -2147217299   ' (&h8004106D)
   wbemErrForcedRollback = -2147217298   ' (&h8004106E)
   wbemErrUnsupportedLocale = -2147217297   ' (&h8004106F)
   wbemErrHandleOutOfDate = -2147217296   ' (&h80041070)
   wbemErrConnectionFailed = -2147217295   ' (&h80041071)
   wbemErrInvalidHandleRequest = -2147217294   ' (&h80041072)
   wbemErrPropertyNameTooWide = -2147217293   ' (&h80041073)
   wbemErrClassNameTooWide = -2147217292   ' (&h80041074)
   wbemErrMethodNameTooWide = -2147217291   ' (&h80041075)
   wbemErrQualifierNameTooWide = -2147217290   ' (&h80041076)
   wbemErrRerunCommand = -2147217289   ' (&h80041077)
   wbemErrDatabaseVerMismatch = -2147217288   ' (&h80041078)
   wbemErrVetoPut = -2147217287   ' (&h80041079)
   wbemErrVetoDelete = -2147217286   ' (&h8004107A)
   wbemErrInvalidLocale = -2147217280   ' (&h80041080)
   wbemErrProviderSuspended = -2147217279   ' (&h80041081)
   wbemErrSynchronizationRequired = -2147217278   ' (&h80041082)
   wbemErrNoSchema = -2147217277   ' (&h80041083)
   wbemErrProviderAlreadyRegistered = -2147217276   ' (&h80041084)
   wbemErrProviderNotRegistered = -2147217275   ' (&h80041085)
   wbemErrFatalTransportError = -2147217274   ' (&h80041086)
   wbemErrEncryptedConnectionRequired = -2147217273   ' (&h80041087)
   wbemErrRegistrationTooBroad = -2147213311   ' (&h80042001)
   wbemErrRegistrationTooPrecise = -2147213310   ' (&h80042002)
   wbemErrTimedout = -2147209215   ' (&h80043001)
   wbemErrResetToDefault = -2147209214   ' (&h80043002)
END ENUM

ENUM WbemFlagEnum
   ' // IID: {4A249B73-FC9A-11D1-8B1E-00600806D9B6}
   ' // Documentation string: Defines behavior of various interface calls
   ' // Number of constants: 15
   wbemFlagReturnImmediately = 16   ' (&h00000010)
   wbemFlagReturnWhenComplete = 0   ' (&h00000000)
   wbemFlagBidirectional = 0   ' (&h00000000)
   wbemFlagForwardOnly = 32   ' (&h00000020)
   wbemFlagNoErrorObject = 64   ' (&h00000040)
   wbemFlagReturnErrorObject = 0   ' (&h00000000)
   wbemFlagSendStatus = 128   ' (&h00000080)
   wbemFlagDontSendStatus = 0   ' (&h00000000)
   wbemFlagEnsureLocatable = 256   ' (&h00000100)
   wbemFlagDirectRead = 512   ' (&h00000200)
   wbemFlagSendOnlySelected = 0   ' (&h00000000)
   wbemFlagUseAmendedQualifiers = 131072   ' (&h00020000)
   wbemFlagGetDefault = 0   ' (&h00000000)
   wbemFlagSpawnInstance = 1   ' (&h00000001)
   wbemFlagUseCurrentTime = 1   ' (&h00000001)
END ENUM

ENUM WbemImpersonationLevelEnum
   ' // IID: {B54D66E8-2287-11D2-8B33-00600806D9B6}
   ' // Documentation string: Defines the security impersonation level
   ' // Number of constants: 4
   wbemImpersonationLevelAnonymous = 1   ' (&h00000001)
   wbemImpersonationLevelIdentify = 2   ' (&h00000002)
   wbemImpersonationLevelImpersonate = 3   ' (&h00000003)
   wbemImpersonationLevelDelegate = 4   ' (&h00000004)
END ENUM

ENUM WbemObjectTextFormatEnum
   ' // IID: {09FF1992-EA0E-11D3-B391-00105A1F473A}
   ' // Documentation string: Defines object text formats
   ' // Number of constants: 2
   wbemObjectTextFormatCIMDTD20 = 1   ' (&h00000001)
   wbemObjectTextFormatWMIDTD20 = 2   ' (&h00000002)
END ENUM

ENUM WbemPrivilegeEnum
   ' // IID: {176D2F70-5AF3-11D2-8B4A-00600806D9B6}
   ' // Documentation string: Defines a privilege
   ' // Number of constants: 27
   wbemPrivilegeCreateToken = 1   ' (&h00000001)
   wbemPrivilegePrimaryToken = 2   ' (&h00000002)
   wbemPrivilegeLockMemory = 3   ' (&h00000003)
   wbemPrivilegeIncreaseQuota = 4   ' (&h00000004)
   wbemPrivilegeMachineAccount = 5   ' (&h00000005)
   wbemPrivilegeTcb = 6   ' (&h00000006)
   wbemPrivilegeSecurity = 7   ' (&h00000007)
   wbemPrivilegeTakeOwnership = 8   ' (&h00000008)
   wbemPrivilegeLoadDriver = 9   ' (&h00000009)
   wbemPrivilegeSystemProfile = 10   ' (&h0000000A)
   wbemPrivilegeSystemtime = 11   ' (&h0000000B)
   wbemPrivilegeProfileSingleProcess = 12   ' (&h0000000C)
   wbemPrivilegeIncreaseBasePriority = 13   ' (&h0000000D)
   wbemPrivilegeCreatePagefile = 14   ' (&h0000000E)
   wbemPrivilegeCreatePermanent = 15   ' (&h0000000F)
   wbemPrivilegeBackup = 16   ' (&h00000010)
   wbemPrivilegeRestore = 17   ' (&h00000011)
   wbemPrivilegeShutdown = 18   ' (&h00000012)
   wbemPrivilegeDebug = 19   ' (&h00000013)
   wbemPrivilegeAudit = 20   ' (&h00000014)
   wbemPrivilegeSystemEnvironment = 21   ' (&h00000015)
   wbemPrivilegeChangeNotify = 22   ' (&h00000016)
   wbemPrivilegeRemoteShutdown = 23   ' (&h00000017)
   wbemPrivilegeUndock = 24   ' (&h00000018)
   wbemPrivilegeSyncAgent = 25   ' (&h00000019)
   wbemPrivilegeEnableDelegation = 26   ' (&h0000001A)
   wbemPrivilegeManageVolume = 27   ' (&h0000001B)
END ENUM

ENUM WbemQueryFlagEnum
   ' // IID: {4A249B76-FC9A-11D1-8B1E-00600806D9B6}
   ' // Documentation string: Defines depth of enumeration or query
   ' // Number of constants: 3
   wbemQueryFlagDeep = 0   ' (&h00000000)
   wbemQueryFlagShallow = 1   ' (&h00000001)
   wbemQueryFlagPrototype = 2   ' (&h00000002)
END ENUM

ENUM WbemTextFlagEnum
   ' // IID: {4A249B78-FC9A-11D1-8B1E-00600806D9B6}
   ' // Documentation string: Defines content of generated object text
   ' // Number of constants: 1
   wbemTextFlagNoFlavors = 1   ' (&h00000001)
END ENUM

ENUM WbemTimeout
   ' // IID: {BF078C2A-07D9-11D2-8B21-00600806D9B6}
   ' // Documentation string: Defines timeout constants
   ' // Number of constants: 1
   wbemTimeoutInfinite = -1   ' (&hFFFFFFFF)
END ENUM

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
TYPE AFX_ISWbemDateTime AS AFX_ISWbemDateTime_
TYPE AFX_ISWbemEventSource AS AFX_ISWbemEventSource_
'TYPE AFX_ISWbemLastError AS AFX_ISWbemLastError_
#define AFX_ISWbemLastError AS AFX_ISWbemObject_
TYPE AFX_ISWbemLocator AS AFX_ISWbemLocator_
TYPE AFX_ISWbemMethod AS AFX_ISWbemMethod_
TYPE AFX_ISWbemMethodSet AS AFX_ISWbemMethodSet_
TYPE AFX_ISWbemNamedValue AS AFX_ISWbemNamedValue_
TYPE AFX_ISWbemNamedValueSet AS AFX_ISWbemNamedValueSet_
TYPE AFX_ISWbemObject AS AFX_ISWbemObject_
TYPE AFX_ISWbemObjectEx AS AFX_ISWbemObjectEx_
TYPE AFX_ISWbemObjectPath AS AFX_ISWbemObjectPath_
TYPE AFX_ISWbemObjectSet AS AFX_ISWbemObjectSet_
TYPE AFX_ISWbemPrivilege AS AFX_ISWbemPrivilege_
TYPE AFX_ISWbemPrivilegeSet AS AFX_ISWbemPrivilegeSet_
TYPE AFX_ISWbemProperty AS AFX_ISWbemProperty_
TYPE AFX_ISWbemPropertySet AS AFX_ISWbemPropertySet_
TYPE AFX_ISWbemQualifier AS AFX_ISWbemQualifier_
TYPE AFX_ISWbemQualifierSet AS AFX_ISWbemQualifierSet_
TYPE AFX_ISWbemRefreshableItem AS AFX_ISWbemRefreshableItem_
TYPE AFX_ISWbemRefresher AS AFX_ISWbemRefresher_
TYPE AFX_ISWbemSecurity AS AFX_ISWbemSecurity_
TYPE AFX_ISWbemServices AS AFX_ISWbemServices_
TYPE AFX_ISWbemServicesEx AS AFX_ISWbemServicesEx_
TYPE AFX_ISWbemSink AS AFX_ISWbemSink_

' // Dual interfaces

' ########################################################################################
' Interface name: ISWbemDateTime
' IID: {5E97458A-CF77-11D3-B38F-00105A1F473A}
' Documentation string: A Datetime
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 40
' ########################################################################################

#ifndef __Afx_ISWbemDateTime_INTERFACE_DEFINED__
#define __Afx_ISWbemDateTime_INTERFACE_DEFINED__

TYPE Afx_ISWbemDateTime_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Value (BYVAL strValue AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Value (BYVAL strValue AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Year (BYVAL iYear AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Year (BYVAL iYear AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_YearSpecified (BYVAL bYearSpecified AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_YearSpecified (BYVAL bYearSpecified AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Month (BYVAL iMonth AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Month (BYVAL iMonth AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MonthSpecified (BYVAL bMonthSpecified AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MonthSpecified (BYVAL bMonthSpecified AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Day (BYVAL iDay AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Day (BYVAL iDay AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DaySpecified (BYVAL bDaySpecified AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DaySpecified (BYVAL bDaySpecified AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Hours (BYVAL iHours AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Hours (BYVAL iHours AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HoursSpecified (BYVAL bHoursSpecified AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_HoursSpecified (BYVAL bHoursSpecified AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Minutes (BYVAL iMinutes AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Minutes (BYVAL iMinutes AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MinutesSpecified (BYVAL bMinutesSpecified AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MinutesSpecified (BYVAL bMinutesSpecified AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Seconds (BYVAL iSeconds AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Seconds (BYVAL iSeconds AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SecondsSpecified (BYVAL bSecondsSpecified AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SecondsSpecified (BYVAL bSecondsSpecified AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Microseconds (BYVAL iMicroseconds AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Microseconds (BYVAL iMicroseconds AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MicrosecondsSpecified (BYVAL bMicrosecondsSpecified AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MicrosecondsSpecified (BYVAL bMicrosecondsSpecified AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_UTC (BYVAL iUTC AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_UTC (BYVAL iUTC AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_UTCSpecified (BYVAL bUTCSpecified AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_UTCSpecified (BYVAL bUTCSpecified AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsInterval (BYVAL bIsInterval AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsInterval (BYVAL bIsInterval AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetVarDate (BYVAL bIsLocal AS VARIANT_BOOL = -1, BYVAL dVarDate AS DATE_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetVarDate (BYVAL dVarDate AS DATE_, BYVAL bIsLocal AS VARIANT_BOOL = -1) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFileTime (BYVAL bIsLocal AS VARIANT_BOOL = -1, BYVAL strFileTime AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetFileTime (BYVAL strFileTime AS AFX_BSTR, BYVAL bIsLocal AS VARIANT_BOOL = -1) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemEventSource
' IID: {27D54D92-0EBE-11D2-8B22-00600806D9B6}
' Documentation string: An Event source
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 2
' ########################################################################################

#ifndef __Afx_ISWbemEventSource_INTERFACE_DEFINED__
#define __Afx_ISWbemEventSource_INTERFACE_DEFINED__

TYPE Afx_ISWbemEventSource_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION NextEvent (BYVAL iTimeoutMs AS LONG = -1, BYVAL objWbemObject AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Security_ (BYVAL objWbemSecurity AS Afx_ISWbemSecurity PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemLastError
' IID: {D962DB84-D4BB-11D1-8B09-00600806D9B6}
' Documentation string: The last error on the current thread
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = ISWbemObject
' Number of methods = 0
' ########################################################################################

'#ifndef __Afx_ISWbemLastError_INTERFACE_DEFINED__
'#define __Afx_ISWbemLastError_INTERFACE_DEFINED__

'TYPE Afx_ISWbemLastError_ EXTENDS Afx_ISWbemObject
'   ' // No additional methods
'END TYPE

'#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemLocator
' IID: {76A6415B-CB41-11D1-8B02-00600806D9B6}
' Documentation string: Used to obtain Namespace connections
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 2
' ########################################################################################

#ifndef __Afx_ISWbemLocator_INTERFACE_DEFINED__
#define __Afx_ISWbemLocator_INTERFACE_DEFINED__

TYPE Afx_ISWbemLocator_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION ConnectServer (BYVAL strServer AS AFX_BSTR, BYVAL strNamespace AS AFX_BSTR, BYVAL strUser AS AFX_BSTR, BYVAL strPassword AS AFX_BSTR, BYVAL strLocale AS AFX_BSTR, BYVAL strAuthority AS AFX_BSTR, BYVAL iSecurityFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemServices AS Afx_ISWbemServices PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Security_ (BYVAL objWbemSecurity AS Afx_ISWbemSecurity PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemMethod
' IID: {422E8E90-D955-11D1-8B09-00600806D9B6}
' Documentation string: A Method
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 5
' ########################################################################################

#ifndef __Afx_ISWbemMethod_INTERFACE_DEFINED__
#define __Afx_ISWbemMethod_INTERFACE_DEFINED__

TYPE Afx_ISWbemMethod_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL strName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Origin (BYVAL strOrigin AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_InParameters (BYVAL objWbemInParameters AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_OutParameters (BYVAL objWbemOutParameters AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Qualifiers_ (BYVAL objWbemQualifierSet AS Afx_ISWbemQualifierSet PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemMethodSet
' IID: {C93BA292-D955-11D1-8B09-00600806D9B6}
' Documentation string: A collection of Methods
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 3
' ########################################################################################

#ifndef __Afx_ISWbemMethodSet_INTERFACE_DEFINED__
#define __Afx_ISWbemMethodSet_INTERFACE_DEFINED__

TYPE Afx_ISWbemMethodSet_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL pUnk AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Item (BYVAL strName AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemMethod AS Afx_ISWbemMethod PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL iCount AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemNamedValue
' IID: {76A64164-CB41-11D1-8B02-00600806D9B6}
' Documentation string: A named value
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 3
' ########################################################################################

#ifndef __Afx_ISWbemNamedValue_INTERFACE_DEFINED__
#define __Afx_ISWbemNamedValue_INTERFACE_DEFINED__

TYPE Afx_ISWbemNamedValue_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Value (BYVAL varValue AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Value (BYVAL varValue AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL strName AS AFX_BSTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemNamedValueSet
' IID: {CF2376EA-CE8C-11D1-8B05-00600806D9B6}
' Documentation string: A collection of named values
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 7
' ########################################################################################

#ifndef __Afx_ISWbemNamedValueSet_INTERFACE_DEFINED__
#define __Afx_ISWbemNamedValueSet_INTERFACE_DEFINED__

TYPE Afx_ISWbemNamedValueSet_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL pUnk AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Item (BYVAL strName AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValue AS Afx_ISWbemNamedValue PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL iCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Add (BYVAL strName AS AFX_BSTR, BYVAL varValue AS VARIANT PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValue AS Afx_ISWbemNamedValue PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL strName AS AFX_BSTR, BYVAL iFlags AS LONG = 0) AS HRESULT
   DECLARE ABSTRACT FUNCTION Clone (BYVAL objWbemNamedValueSet AS Afx_ISWbemNamedValueSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteAll () AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemObject
' IID: {76A6415A-CB41-11D1-8B02-00600806D9B6}
' Documentation string: A Class or Instance
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 25
' ########################################################################################

#ifndef __Afx_ISWbemObject_INTERFACE_DEFINED__
#define __Afx_ISWbemObject_INTERFACE_DEFINED__

TYPE Afx_ISWbemObject_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION Put_ (BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectPath AS Afx_ISWbemObjectPath PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION PutAsync_ (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Delete_ (BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteAsync_ (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Instances_ (BYVAL iFlags AS LONG = 16, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION InstancesAsync_ (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Subclasses_ (BYVAL iFlags AS LONG = 16, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SubclassesAsync_ (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Associators_ (BYVAL strAssocClass AS AFX_BSTR, BYVAL strResultClass AS AFX_BSTR, BYVAL strResultRole AS AFX_BSTR, BYVAL strRole AS AFX_BSTR, BYVAL bClassesOnly AS VARIANT_BOOL = 0, BYVAL bSchemaOnly AS VARIANT_BOOL = 0, BYVAL strRequiredAssocQualifier AS AFX_BSTR, BYVAL strRequiredQualifier AS AFX_BSTR, BYVAL iFlags AS LONG = 16, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AssociatorsAsync_ (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strAssocClass AS AFX_BSTR, BYVAL strResultClass AS AFX_BSTR, BYVAL strResultRole AS AFX_BSTR, BYVAL strRole AS AFX_BSTR, BYVAL bClassesOnly AS VARIANT_BOOL = 0, BYVAL bSchemaOnly AS VARIANT_BOOL = 0, BYVAL strRequiredAssocQualifier AS AFX_BSTR, BYVAL strRequiredQualifier AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION References_ (BYVAL strResultClass AS AFX_BSTR, BYVAL strRole AS AFX_BSTR, BYVAL bClassesOnly AS VARIANT_BOOL = 0, BYVAL bSchemaOnly AS VARIANT_BOOL = 0, BYVAL strRequiredQualifier AS AFX_BSTR, BYVAL iFlags AS LONG = 16, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReferencesAsync_ (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strResultClass AS AFX_BSTR, BYVAL strRole AS AFX_BSTR, BYVAL bClassesOnly AS VARIANT_BOOL = 0, BYVAL bSchemaOnly AS VARIANT_BOOL = 0, BYVAL strRequiredQualifier AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecMethod_ (BYVAL strMethodName AS AFX_BSTR, BYVAL objWbemInParameters AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemOutParameters AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecMethodAsync_ (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strMethodName AS AFX_BSTR, BYVAL objWbemInParameters AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Clone_ (BYVAL objWbemObject AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetObjectText_ (BYVAL iFlags AS LONG = 0, BYVAL strObjectText AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SpawnDerivedClass_ (BYVAL iFlags AS LONG = 0, BYVAL objWbemObject AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SpawnInstance_ (BYVAL iFlags AS LONG = 0, BYVAL objWbemObject AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CompareTo_ (BYVAL objWbemObject AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, BYVAL bResult AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Qualifiers_ (BYVAL objWbemQualifierSet AS Afx_ISWbemQualifierSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Properties_ (BYVAL objWbemPropertySet AS Afx_ISWbemPropertySet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Methods_ (BYVAL objWbemMethodSet AS Afx_ISWbemMethodSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Derivation_ (BYVAL strClassNameArray AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Path_ (BYVAL objWbemObjectPath AS Afx_ISWbemObjectPath PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Security_ (BYVAL objWbemSecurity AS Afx_ISWbemSecurity PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemObjectEx
' IID: {269AD56A-8A67-4129-BC8C-0506DCFE9880}
' Documentation string: A Class or Instance
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = ISWbemObject
' Number of methods = 4
' ########################################################################################

#ifndef __Afx_ISWbemObjectEx_INTERFACE_DEFINED__
#define __Afx_ISWbemObjectEx_INTERFACE_DEFINED__

TYPE Afx_ISWbemObjectEx_ EXTENDS Afx_ISWbemObject
   DECLARE ABSTRACT FUNCTION Refresh_ (BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SystemProperties_ (BYVAL objWbemPropertySet AS Afx_ISWbemPropertySet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetText_ (BYVAL iObjectTextFormat AS WbemObjectTextFormatEnum, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL bsText AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetFromText_ (BYVAL bsText AS AFX_BSTR, BYVAL iObjectTextFormat AS WbemObjectTextFormatEnum, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemObjectPath
' IID: {5791BC27-CE9C-11D1-97BF-0000F81E849C}
' Documentation string: An Object path
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 23
' ########################################################################################

#ifndef __Afx_ISWbemObjectPath_INTERFACE_DEFINED__
#define __Afx_ISWbemObjectPath_INTERFACE_DEFINED__

TYPE Afx_ISWbemObjectPath_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Path (BYVAL strPath AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Path (BYVAL strPath AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RelPath (BYVAL strRelPath AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_RelPath (BYVAL strRelPath AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Server (BYVAL strServer AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Server (BYVAL strServer AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Namespace (BYVAL strNamespace AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Namespace (BYVAL strNamespace AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ParentNamespace (BYVAL strParentNamespace AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DisplayName (BYVAL strDisplayName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DisplayName (BYVAL strDisplayName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Class (BYVAL strClass AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Class (BYVAL strClass AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsClass (BYVAL bIsClass AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetAsClass () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsSingleton (BYVAL bIsSingleton AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetAsSingleton () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Keys (BYVAL objWbemNamedValueSet AS Afx_ISWbemNamedValueSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Security_ (BYVAL objWbemSecurity AS Afx_ISWbemSecurity PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Locale (BYVAL strLocale AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Locale (BYVAL strLocale AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Authority (BYVAL strAuthority AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Authority (BYVAL strAuthority AS AFX_BSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemObjectSet
' IID: {76A6415F-CB41-11D1-8B02-00600806D9B6}
' Documentation string: A collection of Classes or Instances
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 5
' ########################################################################################

#ifndef __Afx_ISWbemObjectSet_INTERFACE_DEFINED__
#define __Afx_ISWbemObjectSet_INTERFACE_DEFINED__

TYPE Afx_ISWbemObjectSet_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL pUnk AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Item (BYVAL strObjectPath AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemObject AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL iCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Security_ (BYVAL objWbemSecurity AS Afx_ISWbemSecurity PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ItemIndex (BYVAL lIndex AS LONG, BYVAL objWbemObject AS Afx_ISWbemObject PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemPrivilege
' IID: {26EE67BD-5804-11D2-8B4A-00600806D9B6}
' Documentation string: A Privilege Override
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 5
' ########################################################################################

#ifndef __Afx_ISWbemPrivilege_INTERFACE_DEFINED__
#define __Afx_ISWbemPrivilege_INTERFACE_DEFINED__

TYPE Afx_ISWbemPrivilege_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_IsEnabled (BYVAL bIsEnabled AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsEnabled (BYVAL bIsEnabled AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL strDisplayName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DisplayName (BYVAL strDisplayName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Identifier (BYVAL iPrivilege AS WbemPrivilegeEnum PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemPrivilegeSet
' IID: {26EE67BF-5804-11D2-8B4A-00600806D9B6}
' Documentation string: A collection of Privilege Overrides
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 7
' ########################################################################################

#ifndef __Afx_ISWbemPrivilegeSet_INTERFACE_DEFINED__
#define __Afx_ISWbemPrivilegeSet_INTERFACE_DEFINED__

TYPE Afx_ISWbemPrivilegeSet_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL pUnk AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Item (BYVAL iPrivilege AS WbemPrivilegeEnum, BYVAL objWbemPrivilege AS Afx_ISWbemPrivilege PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL iCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Add (BYVAL iPrivilege AS WbemPrivilegeEnum, BYVAL bIsEnabled AS VARIANT_BOOL = -1, BYVAL objWbemPrivilege AS Afx_ISWbemPrivilege PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL iPrivilege AS WbemPrivilegeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteAll () AS HRESULT
   DECLARE ABSTRACT FUNCTION AddAsString (BYVAL strPrivilege AS AFX_BSTR, BYVAL bIsEnabled AS VARIANT_BOOL = -1, BYVAL objWbemPrivilege AS Afx_ISWbemPrivilege PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemProperty
' IID: {1A388F98-D4BA-11D1-8B09-00600806D9B6}
' Documentation string: A Property
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 8
' ########################################################################################

#ifndef __Afx_ISWbemProperty_INTERFACE_DEFINED__
#define __Afx_ISWbemProperty_INTERFACE_DEFINED__

TYPE Afx_ISWbemProperty_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Value (BYVAL varValue AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Value (BYVAL varValue AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL strName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsLocal (BYVAL bIsLocal AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Origin (BYVAL strOrigin AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CIMType (BYVAL iCimType AS WbemCimtypeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Qualifiers_ (BYVAL objWbemQualifierSet AS Afx_ISWbemQualifierSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsArray (BYVAL bIsArray AS VARIANT_BOOL PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemPropertySet
' IID: {DEA0A7B2-D4BA-11D1-8B09-00600806D9B6}
' Documentation string: A collection of Properties
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 5
' ########################################################################################

#ifndef __Afx_ISWbemPropertySet_INTERFACE_DEFINED__
#define __Afx_ISWbemPropertySet_INTERFACE_DEFINED__

TYPE Afx_ISWbemPropertySet_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL pUnk AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Item (BYVAL strName AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemProperty AS Afx_ISWbemProperty PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL iCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Add (BYVAL strName AS AFX_BSTR, BYVAL iCimType AS WbemCimtypeEnum, BYVAL bIsArray AS VARIANT_BOOL = 0, BYVAL iFlags AS LONG = 0, BYVAL objWbemProperty AS Afx_ISWbemProperty PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL strName AS AFX_BSTR, BYVAL iFlags AS LONG = 0) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemQualifier
' IID: {79B05932-D3B7-11D1-8B06-00600806D9B6}
' Documentation string: A Qualifier
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 11
' ########################################################################################

#ifndef __Afx_ISWbemQualifier_INTERFACE_DEFINED__
#define __Afx_ISWbemQualifier_INTERFACE_DEFINED__

TYPE Afx_ISWbemQualifier_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Value (BYVAL varValue AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Value (BYVAL varValue AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL strName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsLocal (BYVAL bIsLocal AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PropagatesToSubclass (BYVAL bPropagatesToSubclass AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PropagatesToSubclass (BYVAL bPropagatesToSubclass AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PropagatesToInstance (BYVAL bPropagatesToInstance AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PropagatesToInstance (BYVAL bPropagatesToInstance AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsOverridable (BYVAL bIsOverridable AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsOverridable (BYVAL bIsOverridable AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsAmended (BYVAL bIsAmended AS VARIANT_BOOL PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemQualifierSet
' IID: {9B16ED16-D3DF-11D1-8B08-00600806D9B6}
' Documentation string: A collection of Qualifiers
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 5
' ########################################################################################

#ifndef __Afx_ISWbemQualifierSet_INTERFACE_DEFINED__
#define __Afx_ISWbemQualifierSet_INTERFACE_DEFINED__

TYPE ISWbemQualifierSet_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL pUnk AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Item (BYVAL Name AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemQualifier AS Afx_ISWbemQualifier PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL iCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Add (BYVAL strName AS AFX_BSTR, BYVAL varVal AS VARIANT PTR, BYVAL bPropagatesToSubclass AS VARIANT_BOOL = -1, BYVAL bPropagatesToInstance AS VARIANT_BOOL = -1, BYVAL bIsOverridable AS VARIANT_BOOL = -1, BYVAL iFlags AS LONG = 0, BYVAL objWbemQualifier AS Afx_ISWbemQualifier PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL strName AS AFX_BSTR, BYVAL iFlags AS LONG = 0) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemRefreshableItem
' IID: {5AD4BF92-DAAB-11D3-B38F-00105A1F473A}
' Documentation string: A single item in a Refresher
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 6
' ########################################################################################

#ifndef __Afx_ISWbemRefreshableItem_INTERFACE_DEFINED__
#define __Afx_ISWbemRefreshableItem_INTERFACE_DEFINED__

TYPE ISWbemRefreshableItem_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Index (BYVAL iIndex AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Refresher (BYVAL objWbemRefresher AS Afx_ISWbemRefresher PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsSet (BYVAL bIsSet AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Object (BYVAL objWbemObject AS Afx_ISWbemObjectEx PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ObjectSet (BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL iFlags AS LONG = 0) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemRefresher
' IID: {14D8250E-D9C2-11D3-B38F-00105A1F473A}
' Documentation string: A Collection of Refreshable Objects
' Attributes =  4544 [&h000011C0] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 10
' ########################################################################################

#ifndef __Afx_ISWbemRefresher_INTERFACE_DEFINED__
#define __Afx_ISWbemRefresher_INTERFACE_DEFINED__

TYPE ISWbemRefresher_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get__NewEnum (BYVAL pUnk AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Item (BYVAL iIndex AS LONG, BYVAL objWbemRefreshableItem AS Afx_ISWbemRefreshableItem PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL iCount AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Add (BYVAL objWbemServices AS Afx_ISWbemServicesEx PTR, BYVAL bsInstancePath AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemRefreshableItem AS Afx_ISWbemRefreshableItem PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddEnum (BYVAL objWbemServices AS Afx_ISWbemServicesEx PTR, BYVAL bsClassName AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemRefreshableItem AS Afx_ISWbemRefreshableItem PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL iIndex AS LONG, BYVAL iFlags AS LONG = 0) AS HRESULT
   DECLARE ABSTRACT FUNCTION Refresh (BYVAL iFlags AS LONG = 0) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AutoReconnect (BYVAL bCount AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AutoReconnect (BYVAL bCount AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteAll () AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemSecurity
' IID: {B54D66E6-2287-11D2-8B33-00600806D9B6}
' Documentation string: A Security Configurator
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 5
' ########################################################################################

#ifndef __Afx_ISWbemSecurity_INTERFACE_DEFINED__
#define __Afx_ISWbemSecurity_INTERFACE_DEFINED__

TYPE ISWbemSecurity_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_ImpersonationLevel (BYVAL iImpersonationLevel AS WbemImpersonationLevelEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ImpersonationLevel (BYVAL iImpersonationLevel AS WbemImpersonationLevelEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AuthenticationLevel (BYVAL iAuthenticationLevel AS WbemAuthenticationLevelEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AuthenticationLevel (BYVAL iAuthenticationLevel AS WbemAuthenticationLevelEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Privileges (BYVAL objWbemPrivilegeSet AS Afx_ISWbemPrivilegeSet PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemServices
' IID: {76A6415C-CB41-11D1-8B02-00600806D9B6}
' Documentation string: A connection to a Namespace
' Attributes =  4432 [&h00001150] [Hidden] [Dual] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 19
' ########################################################################################

#ifndef __Afx_ISWbemServices_INTERFACE_DEFINED__
#define __Afx_ISWbemServices_INTERFACE_DEFINED__

TYPE Afx_ISWbemServices_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION Get (BYVAL strObjectPath AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObject AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAsync (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strObjectPath AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Delete_ (BYVAL strObjectPath AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteAsync (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strObjectPath AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION InstancesOf (BYVAL strClass AS AFX_BSTR, BYVAL iFlags AS LONG = 16, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION InstancesOfAsync (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strClass AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SubclassesOf (BYVAL strSuperclass AS AFX_BSTR, BYVAL iFlags AS LONG = 16, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SubclassesOfAsync (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strSuperclass AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecQuery (BYVAL strQuery AS AFX_BSTR, BYVAL strQueryLanguage AS AFX_BSTR, BYVAL iFlags AS LONG = 16, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecQueryAsync (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strQuery AS AFX_BSTR, BYVAL strQueryLanguage AS AFX_BSTR, BYVAL lFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AssociatorsOf (BYVAL strObjectPath AS AFX_BSTR, BYVAL strAssocClass AS AFX_BSTR, BYVAL strResultClass AS AFX_BSTR, BYVAL strResultRole AS AFX_BSTR, BYVAL strRole AS AFX_BSTR, BYVAL bClassesOnly AS VARIANT_BOOL = 0, BYVAL bSchemaOnly AS VARIANT_BOOL = 0, BYVAL strRequiredAssocQualifier AS AFX_BSTR, BYVAL strRequiredQualifier AS AFX_BSTR, BYVAL iFlags AS LONG = 16, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AssociatorsOfAsync (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strObjectPath AS AFX_BSTR, BYVAL strAssocClass AS AFX_BSTR, BYVAL strResultClass AS AFX_BSTR, BYVAL strResultRole AS AFX_BSTR, BYVAL strRole AS AFX_BSTR, BYVAL bClassesOnly AS VARIANT_BOOL = 0, BYVAL bSchemaOnly AS VARIANT_BOOL = 0, BYVAL strRequiredAssocQualifier AS AFX_BSTR, BYVAL strRequiredQualifier AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReferencesTo (BYVAL strObjectPath AS AFX_BSTR, BYVAL strResultClass AS AFX_BSTR, BYVAL strRole AS AFX_BSTR, BYVAL bClassesOnly AS VARIANT_BOOL = 0, BYVAL bSchemaOnly AS VARIANT_BOOL = 0, BYVAL strRequiredQualifier AS AFX_BSTR, BYVAL iFlags AS LONG = 16, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectSet AS Afx_ISWbemObjectSet PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReferencesToAsync (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strObjectPath AS AFX_BSTR, BYVAL strResultClass AS AFX_BSTR, BYVAL strRole AS AFX_BSTR, BYVAL bClassesOnly AS VARIANT_BOOL = 0, BYVAL bSchemaOnly AS VARIANT_BOOL = 0, BYVAL strRequiredQualifier AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecNotificationQuery (BYVAL strQuery AS AFX_BSTR, BYVAL strQueryLanguage AS AFX_BSTR, BYVAL iFlags AS LONG = 48, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemEventSource AS Afx_ISWbemEventSource PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecNotificationQueryAsync (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strQuery AS AFX_BSTR, BYVAL strQueryLanguage AS AFX_BSTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecMethod (BYVAL strObjectPath AS AFX_BSTR, BYVAL strMethodName AS AFX_BSTR, BYVAL objWbemInParameters AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemOutParameters AS Afx_ISWbemObject PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecMethodAsync (BYVAL objWbemSink AS Afx_IDispatch PTR, BYVAL strObjectPath AS AFX_BSTR, BYVAL strMethodName AS AFX_BSTR, BYVAL objWbemInParameters AS Afx_IDispatch PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Security_ (BYVAL objWbemSecurity AS Afx_ISWbemSecurity PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemServicesEx
' IID: {D2F68443-85DC-427E-91D8-366554CC754C}
' Documentation string: A connection to a Namespace
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = ISWbemServices
' Number of methods = 2
' ########################################################################################

#ifndef __Afx_ISWbemServicesEx_INTERFACE_DEFINED__
#define __Afx_ISWbemServicesEx_INTERFACE_DEFINED__

TYPE Afx_ISWbemServicesEx_ EXTENDS Afx_ISWbemServices
   DECLARE ABSTRACT FUNCTION Put (BYVAL objWbemObject AS Afx_ISWbemObjectEx PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemObjectPath AS Afx_ISWbemObjectPath PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION PutAsync (BYVAL objWbemSink AS Afx_ISWbemSink PTR, BYVAL objWbemObject AS Afx_ISWbemObjectEx PTR, BYVAL iFlags AS LONG = 0, BYVAL objWbemNamedValueSet AS Afx_IDispatch PTR, BYVAL objWbemAsyncContext AS Afx_IDispatch PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ISWbemSink
' IID: {75718C9F-F029-11D1-A1AC-00C04FB6C223}
' Documentation string: Asynchronous operation control
' Attributes =  4560 [&h000011D0] [Hidden] [Dual] [Nonextensible] [Oleautomation] [Dispatchable]
' Inherited interface = IDispatch
' Number of methods = 1
' ########################################################################################

#ifndef __Afx_ISWbemSink_INTERFACE_DEFINED__
#define __Afx_ISWbemSink_INTERFACE_DEFINED__

TYPE Afx_ISWbemSink_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION Cancel () AS HRESULT
END TYPE

#endif

' ########################################################################################

END NAMESPACE
