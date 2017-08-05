' ########################################################################################
' Microsoft Windows
' File: AfxADO.bi
' Version: 6.1, Locale ID = 0
' Description: Microsoft ActiveX Data Objects 6.1 Library
' Path: C:\Program Files (x86)\Common Files\System\ado\msado15.dll
' Library GUID: {B691E011-1797-432E-907A-4D8C69339129}
' Portions Copyright (c) Microsoft Corporation.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "win/unknwnbase.bi"

NAMESPACE Afx

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

' ========================================================================================
' Macro for debug
' To allow debugging, define _CADODB_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CADODB_DEBUG_
   #define _CADODB_DEBUG_ 0
#ENDIF
#ifndef _CADODB_DP_
   #define _CADODB_DP_ 1
   #MACRO CADODB_DP(st)
      #IF (_CADODB_DEBUG_ = 1)
         OutputDebugStringW(st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' ProgIDs (Program identifiers)
' ========================================================================================

' CLSID = {00000507-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Command60 = "ADODB.Command.6.0"
' CLSID = {00000514-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Connection60 = "ADODB.Connection.6.0"
' CLSID = {0000050B-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Parameter60 = "ADODB.Parameter.6.0"
' CLSID = {00000560-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Record60 = "ADODB.Record.6.0"
' CLSID = {00000535-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Recordset60 = "ADODB.Recordset.6.0"
' CLSID = {00000566-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Stream60 = "ADODB.Stream.6.0"

' ========================================================================================
' Version independent ProgIDs
' ========================================================================================

' CLSID = {00000507-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Command = "ADODB.Command"
' CLSID = {00000514-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Connection = "ADODB.Connection"
' CLSID = {0000050B-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Parameter = "ADODB.Parameter"
' CLSID = {00000560-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Record = "ADODB.Record"
' CLSID = {00000535-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Recordset = "ADODB.Recordset"
' CLSID = {00000566-0000-0010-8000-00AA006D2EA4}
CONST AFX_PROGID_Stream = "ADODB.Stream"

' ========================================================================================
' ClsIDs (Class identifiers)
' ========================================================================================

CONST AFX_CLSID_Command = "{00000507-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Connection = "{00000514-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Parameter = "{0000050B-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Record = "{00000560-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Recordset = "{00000535-0000-0010-8000-00AA006D2EA4}"
CONST AFX_CLSID_Stream = "{00000566-0000-0010-8000-00AA006D2EA4}"

' ========================================================================================
' IIDs (Interface identifiers)
' ========================================================================================

CONST AFX_IID_ADO = "{00000534-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Collection = "{00000512-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Command = "{986761E8-7269-4890-AA65-AD7C03697A6D}"
CONST AFX_IID_Connection = "{00001550-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_DynaCollection = "{00000513-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Parameter = "{0000150C-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Record = "{00001562-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Recordset = "{00001556-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Stream = "{00001565-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADOCommandConstruction = "{00000517-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADOConnectionConstruction = "{00000551-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADORecordConstruction = "{00000567-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADORecordsetConstruction = "{00000283-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ADOStreamConstruction = "{00000568-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ConnectionEvents = "{00001400-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_ConnectionEventsVt = "{00001402-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Error = "{00000500-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Errors = "{00000501-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Field = "{00001569-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Fields = "{00001564-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Fields_Deprecated = "{00000564-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Parameters = "{0000150D-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Properties = "{00000504-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_Property = "{00000503-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_RecordsetEvents = "{00001266-0000-0010-8000-00AA006D2EA4}"
CONST AFX_IID_RecordsetEventsVt = "{00001403-0000-0010-8000-00AA006D2EA4}"

' // Typedefs
TYPE ADO_LONGPTR AS LONGINT
TYPE PositionEnum_Param AS LONGINT

' // Aliases
TYPE SearchDirection AS SearchDirectionEnum

TYPE CursorTypeEnum AS LONG
enum
   adOpenUnspecified = -1
   adOpenForwardOnly = 0
   adOpenKeyset = 1
   adOpenDynamic = 2
   adOpenStatic = 3
end enum

TYPE CursorOptionEnum AS LONG
enum
   adHoldRecords = &h100
   adMovePrevious = &h200
   adAddNew = &h1000400
   adDelete = &h1000800
   adUpdate = &h1008000
   adBookmark = &h2000
   adApproxPosition = &h4000
   adUpdateBatch = &h10000
   adResync = &h20000
   adNotify  = &h40000
   adFind = &h80000
   adSeek = &h400000
   adIndex = &h800000
end enum

TYPE LockTypeEnum AS LONG
enum
   adLockUnspecified = -1
   adLockReadOnly = 1
   adLockPessimistic = 2
   adLockOptimistic = 3
   adLockBatchOptimistic = 4
end enum

TYPE ExecuteOptionEnum AS LONG
enum
   adOptionUnspecified = -1
   adAsyncExecute = &h10
   adAsyncFetch = &h20
   adAsyncFetchNonBlocking = &h40
   adExecuteNoRecords = &h80
   adExecuteStream = &h400
   adExecuteRecord = &h800
end enum

TYPE ConnectOptionEnum AS LONG
enum
   adConnectUnspecified = -1
   adAsyncConnect = &h10
end enum

TYPE ObjectStateEnum AS LONG
enum
   adStateClosed = 0
   adStateOpen = &h1
   adStateConnecting = &h2
   adStateExecuting = &h4
   adStateFetching = &h8
end enum

TYPE CursorLocationEnum AS LONG
enum
   adUseNone = 1
   adUseServer = 2
   adUseClient = 3
   adUseClientBatch = 3
end enum

TYPE DataTypeEnum AS LONG
enum
   adEmpty = 0
   adTinyInt = 16
   adSmallInt = 2
   adInteger = 3
   adBigInt = 20
   adUnsignedTinyInt = 17
   adUnsignedSmallInt = 18
   adUnsignedInt = 19
   adUnsignedBigInt = 21
   adSingle = 4
   adDouble = 5
   adCurrency = 6
   adDecimal = 14
   adNumeric = 131
   adBoolean = 11
   adError = 10
   adUserDefined = 132
   adVariant = 12
   adIDispatch = 9
   adIUnknown = 13
   adGUID = 72
   adDate = 7
   adDBDate = 133
   adDBTime = 134
   adDBTimeStamp = 135
   adBSTR = 8
   adChar = 129
   adVarChar = 200
   adLongVarChar = 201
   adWChar = 130
   adVarWChar = 202
   adLongVarWChar = 203
   adBinary = 128
   adVarBinary = 204
   adLongVarBinary = 205
   adChapter = 136
   adFileTime = 64
   adPropVariant = 138
   adVarNumeric = 139
   adArray = &h2000
end enum

TYPE FieldAttributeEnum AS LONG
enum
   adFldUnspecified = -1
   adFldMayDefer = &h2
   adFldUpdatable = &h4
   adFldUnknownUpdatable = &h8
   adFldFixed = &h10
   adFldIsNullable = &h20
   adFldMayBeNull = &h40
   adFldLong = &h80
   adFldRowID = &h100
   adFldRowVersion = &h200
   adFldCacheDeferred = &h1000
   adFldIsChapter = &h2000
   adFldNegativeScale = &h4000
   adFldKeyColumn = &h8000
   adFldIsRowURL = &h10000
   adFldIsDefaultStream = &h20000
   adFldIsCollection = &h40000
end enum

TYPE EditModeEnum AS LONG
enum
   adEditNone = 0
   adEditInProgress = &h1
   adEditAdd = &h2
   adEditDelete = &h4
end enum

TYPE RecordStatusEnum AS LONG
enum
   adRecOK = 0
   adRecNew = &h1
   adRecModified = &h2
   adRecDeleted = &h4
   adRecUnmodified = &h8
   adRecInvalid = &h10
   adRecMultipleChanges = &h40
   adRecPendingChanges = &h80
   adRecCanceled = &h100
   adRecCantRelease = &h400
   adRecConcurrencyViolation = &h800
   adRecIntegrityViolation = &h1000
   adRecMaxChangesExceeded = &h2000
   adRecObjectOpen = &h4000
   adRecOutOfMemory = &h8000
   adRecPermissionDenied = &h10000
   adRecSchemaViolation = &h20000
   adRecDBDeleted = &h40000
end enum

TYPE GetRowsOptionEnum AS LONG
enum
   adGetRowsRest = -1
end enum

TYPE PositionEnum AS LONG
enum
   adPosUnknown = -1
   adPosBOF = -2
   adPosEOF = -3
end enum

TYPE BookmarkEnum AS LONG
enum
   adBookmarkCurrent = 0
   adBookmarkFirst = 1
   adBookmarkLast = 2
end enum

TYPE MarshalOptionsEnum AS LONG
enum
   adMarshalAll = 0
   adMarshalModifiedOnly = 1
end enum

TYPE AffectEnum AS LONG
enum
   adAffectCurrent = 1
   adAffectGroup = 2
   adAffectAll = 3
   adAffectAllChapters = 4
end enum

TYPE ResyncEnum AS LONG
enum
   adResyncUnderlyingValues = 1
   adResyncAllValues = 2
end enum

TYPE CompareEnum AS LONG
enum
   adCompareLessThan = 0
   adCompareEqual = 1
   adCompareGreaterThan = 2
   adCompareNotEqual = 3
   adCompareNotComparable = 4
end enum

TYPE FilterGroupEnum AS LONG
enum
   adFilterNone = 0
   adFilterPendingRecords = 1
   adFilterAffectedRecords = 2
   adFilterFetchedRecords = 3
   adFilterPredicate = 4
   adFilterConflictingRecords = 5
end enum

TYPE SearchDirectionEnum AS LONG
enum
   adSearchForward = 1
   adSearchBackward = -1
end enum

TYPE PersistFormatEnum AS LONG
enum
   adPersistADTG = 0
   adPersistXML = 1
end enum

TYPE StringFormatEnum AS LONG
enum
   adClipString = 2
end enum

TYPE ConnectPromptEnum AS LONG
enum
   adPromptAlways = 1
   adPromptComplete = 2
   adPromptCompleteRequired = 3
   adPromptNever = 4
end enum

TYPE ConnectModeEnum AS LONG
enum
   adModeUnknown = 0
   adModeRead = 1
   adModeWrite = 2
   adModeReadWrite = 3
   adModeShareDenyRead = 4
   adModeShareDenyWrite = 8
   adModeShareExclusive = &hc
   adModeShareDenyNone = &h10
   adModeRecursive = &h400000
end enum

TYPE RecordCreateOptionsEnum AS LONG
enum
   adCreateCollection = &h2000
   adCreateStructDoc = &h80000000
   adCreateNonCollection = 0
   adOpenIfExists = &h2000000
   adCreateOverwrite = &h4000000
   adFailIfNotExists = -1
end enum

TYPE RecordOpenOptionsEnum AS LONG
enum
   adOpenRecordUnspecified = -1
   adOpenSource = &h800000
   adOpenOutput = &h800000
   adOpenAsync = &h1000
   adDelayFetchStream = &h4000
   adDelayFetchFields = &h8000
   adOpenExecuteCommand = &h10000
end enum

TYPE IsolationLevelEnum AS LONG
enum
   adXactUnspecified = &hffffffff
   adXactChaos = &h10
   adXactReadUncommitted = &h100
   adXactBrowse = &h100
   adXactCursorStability = &h1000
   adXactReadCommitted = &h1000
   adXactRepeatableRead = &h10000
   adXactSerializable = &h100000
   adXactIsolated = &h100000
end enum

TYPE XactAttributeEnum AS LONG
enum
   adXactCommitRetaining = &h20000
   adXactAbortRetaining = &h40000
   adXactAsyncPhaseOne = &h80000
   adXactSyncPhaseOne = &h100000
end enum

TYPE PropertyAttributesEnum AS LONG
enum
   adPropNotSupported = 0
   adPropRequired = &h1
   adPropOptional = &h2
   adPropRead = &h200
   adPropWrite = &h400
end enum

TYPE ErrorValueEnum AS LONG
enum
   adErrProviderFailed = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hbb8)
   adErrInvalidArgument = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hbb9)
   adErrOpeningFile = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hbba)
   adErrReadFile = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hbbb)
   adErrWriteFile = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hbbc)
   adErrNoCurrentRecord = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hbcd)
   adErrIllegalOperation = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hc93)
   adErrCantChangeProvider = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hc94)
   adErrInTransaction = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hcae)
   adErrFeatureNotAvailable = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hcb3)
   adErrItemNotFound = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hcc1)
   adErrObjectInCollection = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hd27)
   adErrObjectNotSet = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hd5c)
   adErrDataConversion = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hd5d)
   adErrObjectClosed = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he78)
   adErrObjectOpen = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he79)
   adErrProviderNotFound = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he7a)
   adErrBoundToCommand = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he7b)
   adErrInvalidParamInfo = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he7c)
   adErrInvalidConnection = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he7d)
   adErrNotReentrant = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he7e)
   adErrStillExecuting = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he7f)
   adErrOperationCancelled = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he80)
   adErrStillConnecting = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he81)
   adErrInvalidTransaction = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he82)
   adErrNotExecuting = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he83)
   adErrUnsafeOperation = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he84)
   adwrnSecurityDialog = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he85)
   adwrnSecurityDialogHeader = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he86)
   adErrIntegrityViolation = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he87)
   adErrPermissionDenied = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he88)
   adErrDataOverflow = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he89)
   adErrSchemaViolation = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he8a)
   adErrSignMismatch = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he8b)
   adErrCantConvertvalue = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he8c)
   adErrCantCreate = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he8d)
   adErrColumnNotOnThisRow = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he8e)
   adErrURLDoesNotExist = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he8f)
   adErrTreePermissionDenied = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he90)
   adErrInvalidURL = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he91)
   adErrResourceLocked = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he92)
   adErrResourceExists = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he93)
   adErrCannotComplete = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he94)
   adErrVolumeNotFound = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he95)
   adErrOutOfSpace = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he96)
   adErrResourceOutOfScope = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he97)
   adErrUnavailable = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he98)
   adErrURLNamedRowDoesNotExist = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he99)
   adErrDelResOutOfScope = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he9a)
   adErrPropInvalidColumn = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he9b)
   adErrPropInvalidOption = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he9c)
   adErrPropInvalidValue = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he9d)
   adErrPropConflicting = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he9e)
   adErrPropNotAllSettable = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &he9f)
   adErrPropNotSet = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hea0)
   adErrPropNotSettable = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hea1)
   adErrPropNotSupported = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hea2)
   adErrCatalogNotSet = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hea3)
   adErrCantChangeConnection = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hea4)
   adErrFieldsUpdateFailed = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hea5)
   adErrDenyNotSupported = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hea6)
   adErrDenyTypeNotSupported = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hea7)
   adErrProviderNotSpecified = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &hea9)
   adErrConnectionStringTooLong = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_CONTROL, &heaa)
end enum

TYPE ParameterAttributesEnum AS LONG
enum
   adParamSigned = &h10
   adParamNullable = &h40
   adParamLong = &h80
end enum

TYPE ParameterDirectionEnum AS LONG
enum
   adParamUnknown = 0
   adParamInput = &h1
   adParamOutput = &h2
   adParamInputOutput = &h3
   adParamReturnValue = &h4
end enum

TYPE CommandTypeEnum AS LONG
enum
   adCmdUnspecified = -1
   adCmdUnknown = &h8
   adCmdText = &h1
   adCmdTable = &h2
   adCmdStoredProc = &h4
   adCmdFile = &h100
   adCmdTableDirect = &h200
end enum

TYPE EventStatusEnum AS LONG
enum
   adStatusOK = &h1
   adStatusErrorsOccurred = &h2
   adStatusCantDeny = &h3
   adStatusCancel = &h4
   adStatusUnwantedEvent = &h5
end enum

TYPE EventReasonEnum AS LONG
enum
   adRsnAddNew = 1
   adRsnDelete = 2
   adRsnUpdate = 3
   adRsnUndoUpdate = 4
   adRsnUndoAddNew = 5
   adRsnUndoDelete = 6
   adRsnRequery = 7
   adRsnResynch = 8
   adRsnClose = 9
   adRsnMove = 10
   adRsnFirstChange = 11
   adRsnMoveFirst = 12
   adRsnMoveNext = 13
   adRsnMovePrevious = 14
   adRsnMoveLast = 15
end enum

TYPE SchemaEnum AS LONG
enum
   adSchemaProviderSpecific = -1
   adSchemaAsserts = 0
   adSchemaCatalogs = 1
   adSchemaCharacterSets = 2
   adSchemaCollations = 3
   adSchemaColumns = 4
   adSchemaCheckConstraints = 5
   adSchemaConstraintColumnUsage = 6
   adSchemaConstraintTableUsage = 7
   adSchemaKeyColumnUsage = 8
   adSchemaReferentialContraints = 9
   adSchemaReferentialConstraints = 9
   adSchemaTableConstraints = 10
   adSchemaColumnsDomainUsage = 11
   adSchemaIndexes = 12
   adSchemaColumnPrivileges = 13
   adSchemaTablePrivileges = 14
   adSchemaUsagePrivileges = 15
   adSchemaProcedures = 16
   adSchemaSchemata = 17
   adSchemaSQLLanguages = 18
   adSchemaStatistics = 19
   adSchemaTables = 20
   adSchemaTranslations = 21
   adSchemaProviderTypes = 22
   adSchemaViews = 23
   adSchemaViewColumnUsage = 24
   adSchemaViewTableUsage = 25
   adSchemaProcedureParameters = 26
   adSchemaForeignKeys = 27
   adSchemaPrimaryKeys = 28
   adSchemaProcedureColumns = 29
   adSchemaDBInfoKeywords = 30
   adSchemaDBInfoLiterals = 31
   adSchemaCubes = 32
   adSchemaDimensions = 33
   adSchemaHierarchies = 34
   adSchemaLevels = 35
   adSchemaMeasures = 36
   adSchemaProperties = 37
   adSchemaMembers = 38
   adSchemaTrustees = 39
   adSchemaFunctions = 40
   adSchemaActions = 41
   adSchemaCommands = 42
   adSchemaSets = 43
end enum

TYPE FieldStatusEnum AS LONG
enum
   adFieldOK = 0
   adFieldCantConvertValue = 2
   adFieldIsNull = 3
   adFieldTruncated = 4
   adFieldSignMismatch = 5
   adFieldDataOverflow = 6
   adFieldCantCreate = 7
   adFieldUnavailable = 8
   adFieldPermissionDenied = 9
   adFieldIntegrityViolation = 10
   adFieldSchemaViolation = 11
   adFieldBadStatus = 12
   adFieldDefault = 13
   adFieldIgnore = 15
   adFieldDoesNotExist = 16
   adFieldInvalidURL = 17
   adFieldResourceLocked = 18
   adFieldResourceExists = 19
   adFieldCannotComplete = 20
   adFieldVolumeNotFound = 21
   adFieldOutOfSpace = 22
   adFieldCannotDeleteSource = 23
   adFieldReadOnly = 24
   adFieldResourceOutOfScope = 25
   adFieldAlreadyExists = 26
   adFieldPendingInsert = &h10000
   adFieldPendingDelete = &h20000
   adFieldPendingChange = &h40000
   adFieldPendingUnknown = &h80000
   adFieldPendingUnknownDelete = &h100000
end enum

TYPE SeekEnum AS LONG
enum
   adSeekFirstEQ = &h1
   adSeekLastEQ = &h2
   adSeekAfterEQ = &h4
   adSeekAfter = &h8
   adSeekBeforeEQ = &h10
   adSeekBefore = &h20
end enum

TYPE ADCPROP_UPDATECRITERIA_ENUM AS LONG
enum
   adCriteriaKey = 0
   adCriteriaAllCols = 1
   adCriteriaUpdCols = 2
   adCriteriaTimeStamp = 3
end enum

TYPE ADCPROP_ASYNCTHREADPRIORITY_ENUM AS LONG
enum
   adPriorityLowest = 1
   adPriorityBelowNormal = 2
   adPriorityNormal = 3
   adPriorityAboveNormal = 4
   adPriorityHighest = 5
end enum

TYPE ADCPROP_AUTORECALC_ENUM AS LONG
enum
   adRecalcUpFront = 0
   adRecalcAlways = 1
end enum

TYPE ADCPROP_UPDATERESYNC_ENUM AS LONG
enum
   adResyncNone = 0
   adResyncAutoIncrement = 1
   adResyncConflicts = 2
   adResyncUpdates = 4
   adResyncInserts = 8
   adResyncAll = 15
end enum

TYPE MoveRecordOptionsEnum AS LONG
enum
   adMoveUnspecified = -1
   adMoveOverWrite = 1
   adMoveDontUpdateLinks = 2
   adMoveAllowEmulation = 4
end enum

TYPE CopyRecordOptionsEnum AS LONG
enum
   adCopyUnspecified = -1
   adCopyOverWrite = 1
   adCopyAllowEmulation = 4
   adCopyNonRecursive = 2
end enum

TYPE StreamTypeEnum AS LONG
enum
   adTypeBinary = 1
   adTypeText = 2
end enum

TYPE LineSeparatorEnum AS LONG
enum
   adLF = 10
   adCR = 13
   adCRLF = -1
end enum

TYPE StreamOpenOptionsEnum AS LONG
enum
   adOpenStreamUnspecified = -1
   adOpenStreamAsync = 1
   adOpenStreamFromRecord = 4
end enum

TYPE StreamWriteEnum AS LONG
enum
   adWriteChar = 0
   adWriteLine = 1
   stWriteChar = 0
   stWriteLine = 1
end enum

TYPE SaveOptionsEnum AS LONG
enum
   adSaveCreateNotExist = 1
   adSaveCreateOverWrite = 2
end enum

TYPE FieldEnum AS LONG
enum
   adDefaultStream = -1
   adRecordURL = -2
end enum

TYPE StreamReadEnum AS LONG
enum
   adReadAll = -1
   adReadLine = -2
end enum

TYPE RecordTypeEnum AS LONG
enum
   adSimpleRecord = 0
   adCollectionRecord = 1
   adStructDoc = 2
end enum

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

' ########################################################################################
' Forward declarations
' ########################################################################################

TYPE Afx_ADO AS Afx_ADO_
TYPE Afx_ADOCollection AS Afx_ADOCollection_
TYPE Afx_ADODynaCollection AS Afx_ADODynaCollection_
TYPE Afx_ADOProperties AS Afx_ADOProperties_
TYPE Afx_ADOProperty AS Afx_ADOProperty_
TYPE Afx_ADOCommand AS Afx_ADOCommand_
TYPE Afx_ADOConnection AS Afx_ADOConnection_
TYPE Afx_ADORecordset AS Afx_ADORecordset_
TYPE Afx_ADOErrors AS Afx_ADOErrors_
TYPE Afx_ADOError AS Afx_ADOError_
TYPE Afx_ADOFields AS Afx_ADOFields_
TYPE Afx_ADOField AS Afx_ADOField_
TYPE Afx_ADOParameters AS Afx_ADOParameters_
TYPE Afx_ADOParameter AS Afx_ADOParameter_
TYPE Afx_ADORecord AS Afx_ADORecord_
TYPE Afx_ADOStream AS Afx_ADOStream_
TYPE Afx_ADOCommandConstruction AS Afx_ADOCommandConstruction_
TYPE Afx_ADORecordsetConstruction AS Afx_ADORecordsetConstruction_
TYPE Afx_ADOConnectionConstruction AS Afx_ADOConnectionConstruction_
TYPE Afx_ADORecordConstruction AS Afx_ADORecordConstruction_
TYPE Afx_ADOStreamConstruction AS Afx_ADOStreamConstruction_
TYPE Afx_ConnectionEventsVt AS Afx_ConnectionEventsVt_
TYPE Afx_RecordsetEventsVt AS Afx_RecordsetEventsVt_

' ########################################################################################
' Interface name = _ADO
' IID = {00000534-0000-0010-8000-00AA006D2EA4}
' Attributes = 4288 [&H000010C0] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADO_INTERFACE_DEFINED__
#define __Afx_ADO_INTERFACE_DEFINED__
TYPE Afx_ADO_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Properties (BYVAL ppvObject AS Afx_ADOProperties PTR PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = _Collection
' IID = {00000512-0000-0010-8000-00AA006D2EA4}
' Attributes = 4288 [&H000010C0] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADOCollection_INTERFACE_DEFINED__
#define __Afx_ADOCollection_INTERFACE_DEFINED__
TYPE Afx_ADOCollection_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL c AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION _NewEnum (BYVAL ppvObject AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Refresh () AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = _DynaCollection
' IID = {00000513-0000-0010-8000-00AA006D2EA4}
' Attributes = 4288 [&H000010C0] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_ADOCollection
' ########################################################################################

#ifndef __Afx_ADODynaCollection_INTERFACE_DEFINED__
#define __Afx_ADODynaCollection_INTERFACE_DEFINED__
TYPE Afx_ADODynaCollection_ EXTENDS Afx_ADOCollection
   DECLARE ABSTRACT FUNCTION Append (BYVAL pObject AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Delete_ (BYVAL Index AS VARIANT) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = Properties
' IID = {00000504-0000-0010-8000-00AA006D2EA4}
' Attributes = 4288 [&H000010C0] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_ADOCollection
' ########################################################################################

#ifndef __Afx_ADOProperties_INTERFACE_DEFINED__
#define __Afx_ADOProperties_INTERFACE_DEFINED__
TYPE Afx_ADOProperties_ EXTENDS Afx_ADOCollection
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL Index AS VARIANT, BYVAL ppvObject AS Afx_AdoProperty PTR PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = Property
' IID = {00000503-0000-0010-8000-00AA006D2EA4}
' Attributes = 4288 [&H000010C0] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADOProperty_INTERFACE_DEFINED__
#define __Afx_ADOProperty_INTERFACE_DEFINED__
TYPE Afx_ADOProperty_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Value (BYVAL pval AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Value (BYVAL val AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Type (BYVAL ptype AS DataTypeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Attributes (BYVAL plAttributes AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Attributes (BYVAL lAttributes AS LONG) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = Errors
' IID = {00000501-0000-0010-8000-00AA006D2EA4}
' Attributes = 4288 [&H000010C0] [Dual] [Nonextensible] [Dispatchable]
' Extends AFX_ADOCollection
' ########################################################################################

#ifndef __Afx_ADOErrors_INTERFACE_DEFINED__
#define __Afx_ADOErrors_INTERFACE_DEFINED__
TYPE Afx_ADOErrors_ EXTENDS Afx_ADOCollection
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL Index AS VARIANT, BYVAL ppvObject AS Afx_ADOError PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Clear () AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = Error
' IID = {00000500-0000-0010-8000-00AA006D2EA4}
' Attributes = 4288 [&H000010C0] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADOError_INTERFACE_DEFINED__
#define __Afx_ADOError_INTERFACE_DEFINED__
TYPE Afx_ADOError_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Number (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Source (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Description (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HelpFile (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HelpContext (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SQLState (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NativeError (BYVAL pl AS LONG PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = _Connection
' IID = {00001550-0000-0010-8000-00AA006D2EA4}
' Attributes = 4176 [&H00001050] [Hidden] [Dual] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADOConnection_INTERFACE_DEFINED__
#define __Afx_ADOConnection_INTERFACE_DEFINED__
TYPE Afx_ADOConnection_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Properties (BYVAL ppvObject AS Afx_ADOProperties PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ConnectionString (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ConnectionString (BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CommandTimeout (BYVAL plTimeout AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CommandTimeout (BYVAL lTimeout AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ConnectionTimeout (BYVAL plTimeout AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ConnectionTimeout (BYVAL lTimeout AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Version (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Close () AS HRESULT
   DECLARE ABSTRACT FUNCTION Execute (BYVAL CommandText AS AFX_BSTR, BYVAL RecordsAffected AS VARIANT PTR, _
           BYVAL Options AS LONG, BYVAL ppiRset AS Afx_ADORecordset PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION BeginTrans (BYVAL TransactionLevel AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CommitTrans () AS HRESULT
   DECLARE ABSTRACT FUNCTION RollbackTrans () AS HRESULT
   DECLARE ABSTRACT FUNCTION Open (BYVAL ConnectionString AS AFX_BSTR = NULL, BYVAL UserID AS AFX_BSTR = NULL, _
           BYVAL Password AS AFX_BSTR = NULL, BYVAL Options AS LONG = adOptionUnspecified) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Errors (BYVAL ppvObject AS Afx_ADOErrors PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DefaultDatabase (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DefaultDatabase (BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsolationLevel (BYVAL Level AS IsolationLevelEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsolationLevel (BYVAL Level AS IsolationLevelEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Attributes (BYVAL plAttr AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Attributes (BYVAL lAttr AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CursorLocation (BYVAL plCursorLoc AS CursorLocationEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Mode (BYVAL plMode AS ConnectModeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Mode (BYVAL lMode AS ConnectModeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Provider (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Provider (BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_State (BYVAL plObjState AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OpenSchema (BYVAL Schema AS SchemaEnum, BYVAL Restrictions AS VARIANT, _
           BYVAL SchemaID AS VARIANT, BYVAL pprset AS Afx_ADORecordset PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Cancel () AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = _Recordset
' IID = {00001556-0000-0010-8000-00AA006D2EA4}
' Attributes = 4304 [&H000010D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADORecordset_INTERFACE_DEFINED__
#define __Afx_ADORecordset_INTERFACE_DEFINED__
TYPE Afx_ADORecordset_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Properties (BYVAL ppvObject AS Afx_ADOProperties PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AbsolutePosition (BYVAL pl AS PositionEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AbsolutePosition (BYVAL Position AS PositionEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION putref_ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ActiveConnection (BYVAL vConn AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ActiveConnection (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BOF (BYVAL pb AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Bookmark (BYVAL pvBookmark AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Bookmark (BYVAL vBookmark AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CacheSize (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CacheSize (BYVAL CacheSize AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CursorType (BYVAL plCursorType AS CursorTypeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CursorType (BYVAL lCursorType AS CursorTypeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_EOF (BYVAL pb AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Fields (BYVAL ppvObject AS Afx_ADOFields PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_LockType (BYVAL plLockType AS LockTypeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_LockType (BYVAL lLockType AS LockTypeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MaxRecords (BYVAL plMaxRecords AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MaxRecords (BYVAL lMaxRecords AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RecordCount (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION putref_Source (BYVAL pcmd AS Afx_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Source (BYVAL bstrConn AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Source (BYVAL pvSource AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddNew (BYVAL FieldList AS VARIANT, BYVAL Values AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION CancelUpdate () AS HRESULT
   DECLARE ABSTRACT FUNCTION Close () AS HRESULT
   DECLARE ABSTRACT FUNCTION Delete_ (BYVAL AffectRecords AS AffectEnum = adAffectCurrent) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetRows (BYVAL Rows AS LONG, BYVAL Start AS VARIANT, BYVAL Fields AS VARIANT, BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Move (BYVAL NumRecords AS LONG, BYVAL Start AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveNext () AS HRESULT
   DECLARE ABSTRACT FUNCTION MovePrevious () AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveFirst () AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveLast () AS HRESULT
   DECLARE ABSTRACT FUNCTION Open (BYVAL Source AS VARIANT, BYVAL ActiveConnection AS VARIANT, _
           BYVAL CursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL LockType AS LockTypeEnum = adLockUnspecified, _
           BYVAL Options AS LONG = adCmdUnspecified) AS HRESULT
   DECLARE ABSTRACT FUNCTION Requery (BYVAL Options AS LONG = adOptionUnspecified) AS HRESULT
   DECLARE ABSTRACT FUNCTION _xResync (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   DECLARE ABSTRACT FUNCTION Update (BYVAL Fields AS VARIANT, BYVAL Values AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AbsolutePage (BYVAL pl AS PositionEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AbsolutePage (BYVAL Page AS PositionEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_EditMode (BYVAL pl AS EditModeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Filter (BYVAL Criteria AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Filter (BYVAL Criteria AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PageCount (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PageSize (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PageSize (BYVAL PageSize AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Sort (BYVAL Criteria AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Sort (BYVAL Criteria AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Status (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_State (BYVAL plObjState AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION _xClone (BYVAL ppvObject AS Afx_ADORecordset PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION UpdateBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   DECLARE ABSTRACT FUNCTION CancelBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CursorLocation (BYVAL plCursorLoc AS CursorLocationEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION NextRecordset (BYVAL RecordsAffected AS VARIANT PTR, BYVAL ppiRs AS Afx_AdoRecordset PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Supports (BYVAL CursorOptions AS CursorOptionEnum, BYVAL pb AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Collect (BYVAL Index AS VARIANT, BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Collect (BYVAL Index AS VARIANT, BYVAL Value AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MarshalOptions (BYVAL peMarshal AS MarshalOptionsEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MarshalOptions (BYVAL eMarshal AS MarshalOptionsEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION Find (BYVAL Criteria AS AFX_BSTR, BYVAL SkipRecords AS LONG, BYVAL SearchDirection AS SearchDirectionEnum, BYVAL Start AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION Cancel () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DataSource (BYVAL ppunkDataSource AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION putref_DataSource (BYVAL punkDataSource AS Afx_IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION _xSave (BYVAL FileName AS AFX_BSTR = NULL, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ActiveCommand (BYVAL ppCmd AS AFX_IDispatch PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_StayInSync (BYVAL bStayInSync AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_StayInSync (BYVAL pbStayInSync AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetString (BYVAL StringFormat AS StringFormatEnum = adClipString, BYVAL NumRows AS LONG = adReadAll, _
           BYVAL ColumnDelimeter AS AFX_BSTR, BYVAL RowDelimeter AS AFX_BSTR, BYVAL NullExpr AS AFX_BSTR, _
           BYVAL pRetString AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DataMember (BYVAL pbstrDataMember AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DataMember (BYVAL bstrDataMember AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CompareBookmarks (BYVAL Bookmark1 AS VARIANT, BYVAL Bookmark2 AS VARIANT, BYVAL pCompare AS CompareEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Clone (BYVAL LockType AS LockTypeEnum = adLockUnspecified, BYVAL ppvObject AS Afx_AdoRecordset PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Resync (BYVAL AffectRecords AS AffectEnum = adAffectAll, BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   DECLARE ABSTRACT FUNCTION Seek (BYVAL KeyValues AS VARIANT, BYVAL SeekOption AS SeekEnum = adSeekFirstEQ) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Index (BYVAL Index AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Index (BYVAL pbstrIndex AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Save (BYVAL Destination AS VARIANT, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = Fields
' IID = {00001564-0000-0010-8000-00AA006D2EA4}
' Attributes = 4304 [&H000010D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_ADOCollection
' ########################################################################################

#ifndef __Afx_ADOFields_INTERFACE_DEFINED__
#define __Afx_ADOFields_INTERFACE_DEFINED__
TYPE Afx_ADOFields_ EXTENDS Afx_ADOCollection
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL Index AS VARIANT, BYVAL ppvObject AS Afx_AdoField PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION _Append (BYVAL Name AS AFX_BSTR, BYVAL Type AS DataTypeEnum, _
           BYVAL DefinedSize AS LONG = 0, BYVAL Attrib AS FieldAttributeEnum = adFldUnspecified) AS HRESULT
   DECLARE ABSTRACT FUNCTION Delete_ (BYVAL Index AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION Append (BYVAL Name AS AFX_BSTR, BYVAL Type AS DataTypeEnum, _
           BYVAL DefinedSize AS LONG = 0, BYVAL Attrib AS FieldAttributeEnum = adFldUnspecified, _
           BYVAL FieldValue AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION Update () AS HRESULT
   DECLARE ABSTRACT FUNCTION Resync (BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   DECLARE ABSTRACT FUNCTION CancelUpdate () AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = Field
' IID = {00001569-0000-0010-8000-00AA006D2EA4}
' Attributes = 4304 [&H000010D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADOField_INTERFACE_DEFINED__
#define __Afx_ADOField_INTERFACE_DEFINED__
TYPE Afx_ADOField_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Properties (BYVAL ppvObject AS Afx_ADOProperties PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ActualSize (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Attributes (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DefinedSize (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Type (BYVAL pDataType AS DataTypeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Value (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Value (BYVAL Val AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Precision (BYVAL pbPrecision AS BYTE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NumericScale (BYVAL pbNumericScale AS BYTE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AppendChunk (BYVAL Data AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetChunk (BYVAL Length AS LONG, BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_OriginalValue (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_UnderlyingValue (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DataFormat (BYVAL ppiDF AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION putref_DataFormat (BYVAL ppiDF AS Afx_IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Precision (BYVAL bPrecision AS BYTE) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_NumericScale (BYVAL bScale AS BYTE) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Type (BYVAL DataType AS DataTypeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DefinedSize (BYVAL lSize AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Attributes (BYVAL lAttributes AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Status (BYVAL pFStatus AS LONG PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = _Command
' IID = {986761E8-7269-4890-AA65-AD7C03697A6D}
' Attributes = 4304 [&H000010D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADOCommand_INTERFACE_DEFINED__
#define __Afx_ADOCommand_INTERFACE_DEFINED__
TYPE Afx_ADOCommand_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Properties (BYVAL ppvObject AS Afx_ADOProperties PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ActiveConnection (BYVAL ppvObject AS Afx_ADOConnection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION putref_ActiveConnection (BYVAL pCon AS Afx_ADOCOnnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ActiveConnection (BYVAL vConn AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CommandText (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CommandText (BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CommandTimeout (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CommandTimeout (BYVAL Timeout AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Prepared (BYVAL pfPrepared AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Prepared (BYVAL fPrepared AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION Execute (BYVAL RecordsAffected AS VARIANT PTR = NULL, BYVAL Parameters AS VARIANT PTR = NULL, _
           BYVAL Options AS LONG = adCmdUnspecified, BYVAL ppirs AS Afx_ADORecordset PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateParameter (BYVAL Name AS AFX_BSTR, BYVAL Type AS DataTypeEnum, _
           BYVAL Direction AS ParameterDirectionEnum, BYVAL Size AS LONG, BYVAL Value AS VARIANT, _
           BYVAL ppiprm AS Afx_ADOParameter PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Parameters (BYVAL ppvObject AS Afx_ADOParameters PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CommandType (BYVAL lCmdType AS CommandTypeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CommandType (BYVAL plCmdType AS CommandTypeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL pbstrName AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Name (BYVAL bstrName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_State (BYVAL plObjState AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Cancel () AS HRESULT
   DECLARE ABSTRACT FUNCTION putref_CommandStream (BYVAL pStream AS Afx_IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CommandStream (BYVAL pvStream AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Dialect (BYVAL bstrDialect AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Dialect (BYVAL pbstrDialect AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_NamedParameters (BYVAL fNamedParameters AS VARIANT_BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NamedParameters (BYVAL pfNamedParameters AS VARIANT_BOOL PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = Parameters
' IID = {0000150D-0000-0010-8000-00AA006D2EA4}
' Attributes = 4288 [&H000010C0] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_ADODynaCollection
' ########################################################################################

#ifndef __Afx_ADOParameters_INTERFACE_DEFINED__
#define __Afx_ADOParameters_INTERFACE_DEFINED__
TYPE Afx_ADOParameters_ EXTENDS Afx_ADODynaCollection
   DECLARE ABSTRACT FUNCTION get_Item (BYVAL Index AS VARIANT, BYVAL ppvObject AS Afx_ADOParameter PTR PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = _Parameter
' IID = {0000150C-0000-0010-8000-00AA006D2EA4}
' Attributes = 4304 [&H000010D0] [Hidden] [Dual] [Nonextensible] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADOParameter_INTERFACE_DEFINED__
#define __Afx_ADOParameter_INTERFACE_DEFINED__
TYPE Afx_ADOParameter_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Properties (BYVAL ppvObject AS Afx_ADOProperties PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Name (BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Name (BYVAL bstr AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Value (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Value (BYVAL val AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Type (BYVAL psDataType AS DataTypeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Type (BYVAL sDataType AS DataTypeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Direction (BYVAL lParmDirection AS ParameterDirectionEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Direction (BYVAL plParmDirection AS ParameterDirectionEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Precision (BYVAL bPrecision AS BYTE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Precision (BYVAL pbPrecision AS BYTE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_NumericScale (BYVAL bScale AS BYTE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NumericScale (BYVAL pbScale AS BYTE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Size (BYVAL l AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Size (BYVAL pl AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AppendChunk (BYVAL Val AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Attributes (BYVAL plParmAttribs AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Attributes (BYVAL lParmAttribs AS LONG) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = _Record
' IID = {00001562-0000-0010-8000-00AA006D2EA4}
' Attributes = 4176 [&H00001050] [Hidden] [Dual] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADORecord_INTERFACE_DEFINED__
#define __Afx_ADORecord_INTERFACE_DEFINED__
TYPE Afx_ADORecord_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Properties (BYVAL ppvObject AS Afx_ADOProperties PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ActiveConnection (BYVAL pvar AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ActiveConnection (BYVAL bstrConn AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION putref_ActiveConnection (BYVAL Con AS AFX_ADOConnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_State (BYVAL pState AS ObjectStateEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Source (BYVAL pStater AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Source (BYVAL Source AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION putref_Source (BYVAL Source AS AFX_IDispatch PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Mode (BYVAL pMode AS ConnectModeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Mode (BYVAL Mode AS ConnectModeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ParentURL (BYVAL pbstrParentURL AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveRecord (BYVAL Source AS AFX_BSTR, BYVAL Destination AS AFX_BSTR, _
           BYVAL UserName AS AFX_BSTR, BYVAL Password AS AFX_BSTR, BYVAL Options AS MoveRecordOptionsEnum, _
           BYVAL Async AS VARIANT_BOOL, BYVAL pbstrNewURL AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CopyRecord (BYVAL Source AS AFX_BSTR, BYVAL Destination AS AFX_BSTR, _
           BYVAL UserName AS AFX_BSTR, BYVAL Password AS AFX_BSTR, BYVAL Options AS CopyRecordOptionsEnum, _
           BYVAL Async AS VARIANT_BOOL, BYVAL pbstrNewURL AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteRecord (BYVAL Source AS AFX_BSTR = NULL, BYVAL Async AS VARIANT_BOOL = 0) AS HRESULT
   DECLARE ABSTRACT FUNCTION Open (BYVAL Source AS VARIANT, BYVAL ActiveConnection AS VARIANT, _
           BYVAL Mode AS ConnectModeEnum = adModeUnknown, BYVAL CreateOptions AS RecordCreateOptionsEnum = adFailIfNotExists, _
           BYVAL UserName AS AFX_BSTR = NULL, BYVAL Password AS AFX_BSTR = NULL) AS HRESULT
   DECLARE ABSTRACT FUNCTION Close () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Fields (BYVAL ppFlds AS Afx_ADOFields PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RecordType (BYVAL pType AS RecordTypeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetChildren (BYVAL ppRSet AS Afx_ADORecordset PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Cancel () AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = _Stream
' IID = {00001565-0000-0010-8000-00AA006D2EA4}
' Attributes = 4176 [&H00001050] [Hidden] [Dual] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADOStream_INTERFACE_DEFINED__
#define __Afx_ADOStream_INTERFACE_DEFINED__
TYPE Afx_ADOStream_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Size (BYVAL pSize AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_EOS (BYVAL pEOS AS VARIANT_BOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Position (BYVAL pPos AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Position (BYVAL Position AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Type (BYVAL pType AS StreamTypeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Type (BYVAL Type AS StreamTypeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_LineSeparator (BYVAL pLS AS LineSeparatorEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_LineSeparator (BYVAL LineSeparator AS LineSeparatorEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_State (BYVAL pState AS ObjectStateEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Mode (BYVAL pMode AS ConnectModeEnum PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Mode (BYVAL Mode AS ConnectModeEnum) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Charset (BYVAL pbstrCharset AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Charset (BYVAL Charset AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Read (BYVAL NumBytes AS LONG, BYVAL pVal AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Open (BYVAL Source AS VARIANT, BYVAL Mode AS ConnectModeEnum = adModeUnknown, _
           BYVAL Options AS StreamOpenOptionsEnum = adOpenStreamUnspecified, BYVAL UserName AS AFX_BSTR, _
           BYVAL Password AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Close () AS HRESULT
   DECLARE ABSTRACT FUNCTION SkipLine () AS HRESULT
   DECLARE ABSTRACT FUNCTION Write (BYVAL Buffer AS VARIANT) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetEOS () AS HRESULT
   DECLARE ABSTRACT FUNCTION CopyTo (BYVAL DestStream AS Afx_ADOStream PTR, BYVAL CharNumber AS LONG = -1) AS HRESULT
   DECLARE ABSTRACT FUNCTION Flush () AS HRESULT
   DECLARE ABSTRACT FUNCTION SaveToFile (BYVAL FileName AS AFX_BSTR, BYVAL Options AS SaveOptionsEnum = adSaveCreateNotExist) AS HRESULT
   DECLARE ABSTRACT FUNCTION LoadFromFile (BYVAL FileName AS AFX_BSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReadText (BYVAL NumChars AS LONG, BYVAL pbstr AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteText (BYVAL Data AS AFX_BSTR, BYVAL Options AS StreamWriteEnum = adWriteChar) AS HRESULT
   DECLARE ABSTRACT FUNCTION Cancel () AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = ADOCommandConstruction
' IID = {00000517-0000-0010-8000-00AA006D2EA4}
' Attributes = 512 [&H00000200] [Restricted]
' Extends Afx_IUnknown
' ########################################################################################

#ifndef __Afx_ADOCommandConstruction_INTERFACE_DEFINED__
#define __Afx_ADOCommandConstruction_INTERFACE_DEFINED__
TYPE Afx_ADOCommandConstruction_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_OLEDBCommand (BYVAL ppOLEDBCommand AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_OLEDBCommand (BYVAL pOLEDBCommand AS Afx_IUnknown PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = ADORecordsetConstruction
' IID = {00000283-0000-0010-8000-00AA006D2EA4}
' Attributes = 4624 [&H00001210] [Hidden] [Restricted] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADORecordsetConstruction_INTERFACE_DEFINED__
#define __Afx_ADORecordsetConstruction_INTERFACE_DEFINED__
TYPE Afx_ADORecordsetConstruction_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Rowset (BYVAL ppRowset AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Rowset (BYVAL pRowset AS Afx_IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Chapter (BYVAL plChapter AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Chapter (BYVAL lChapter AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RowPosition (BYVAL ppRowPos AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_RowPosition (BYVAL pRowPos AS Afx_IUnknown PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = ADOConnectionConstruction
' IID = {00000551-0000-0010-8000-00AA006D2EA4}
' Attributes = 512 [&H00000200] [Restricted]
' Extends Afx_IUnknown
' ########################################################################################

#ifndef __Afx_ADOConnectionConstruction_INTERFACE_DEFINED__
#define __Afx_ADOConnectionConstruction_INTERFACE_DEFINED__
TYPE Afx_ADOConnectionConstruction_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_DSO (BYVAL ppDSO AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Session (BYVAL ppSession AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WrapDSOandSession (BYVAL pDSO AS Afx_IUnknown PTR, BYVAL pSession AS Afx_IUnknown PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = ADORecordConstruction
' IID = {00000567-0000-0010-8000-00AA006D2EA4}
' Attributes = 4608 [&H00001200] [Restricted] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADORecordConstruction_INTERFACE_DEFINED__
#define __Afx_ADORecordConstruction_INTERFACE_DEFINED__
TYPE Afx_ADORecordConstruction_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Row (BYVAL ppRow AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Row (BYVAL pRow AS Afx_IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ParentRow (BYVAL pRow AS Afx_IUnknown PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = ADOStreamConstruction
' IID = {00000568-0000-0010-8000-00AA006D2EA4}
' Attributes = 4608 [&H00001200] [Restricted] [Dispatchable]
' Extends Afx_IDispatch
' ########################################################################################

#ifndef __Afx_ADOStreamConstruction_INTERFACE_DEFINED__
#define __Afx_ADOStreamConstruction_INTERFACE_DEFINED__
TYPE Afx_ADOStreamConstruction_ EXTENDS Afx_IDispatch
   DECLARE ABSTRACT FUNCTION get_Stream (BYVAL ppStm AS Afx_IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Stream (BYVAL pStm AS Afx_IUnknown PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = ConnectionEventsVt
' IID = {00001402-0000-0010-8000-00AA006D2EA4}
' Attributes = 16 [&H00000010] [Hidden]
' Extends Afx_IUnknown
' ########################################################################################

#ifndef __Afx_ConnectionEventsVt_INTERFACE_DEFINED__
#define __Afx_ConnectionEventsVt_INTERFACE_DEFINED__
TYPE Afx_ConnectionEventsVt_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION InfoMessage (BYVAL pError AS Afx_ADOError PTR, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION BeginTransComplete (BYVAL TransactionLevel AS LONG, _
           BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, _
           BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CommitTransComplete (BYVAL pError AS Afx_ADOError PTR, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RollbackTransComplete (BYVAL pError AS Afx_ADOError PTR, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WillExecute (BYVAL Source AS AFX_BSTR PTR, BYVAL CursorType AS CursorTypeEnum PTR, _
           BYVAL LockType AS LockTypeEnum PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, _
           BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, _
           BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecuteComplete (BYVAL RecordsAffected AS LONG, BYVAL pError AS Afx_ADOError PTR, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, _
           BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOCOnnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WillConnect (BYVAL ConnectionString AS AFX_BSTR PTR, BYVAL UserID AS AFX_BSTR PTR, _
           BYVAL Password AS AFX_BSTR PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, _
           BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ConnectComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, _
           BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Disconnect (BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
END TYPE
#endif

' ########################################################################################
' Interface name = RecordsetEventsVt
' IID = {00001403-0000-0010-8000-00AA006D2EA4}
' Attributes = 16 [&H00000010] [Hidden]
' Extends Afx_IUnknown
' ########################################################################################

#ifndef __Afx_RecordsetEventsVt_INTERFACE_DEFINED__
#define __Afx_RecordsetEventsVt_INTERFACE_DEFINED__
TYPE Afx_RecordsetEventsVt_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION WillChangeField (BYVAL cFields AS LONG, BYVAL Fields AS VARIANT, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION FieldChangeComplete (BYVAL cFields AS LONG, BYVAL Fields AS VARIANT, _
           BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, _
           BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WillChangeRecord (BYVAL adReason AS EventReasonEnum, BYVAL cRecords AS LONG, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RecordChangeComplete (BYVAL adReason AS EventReasonEnum, BYVAL cRecords AS LONG, BYVAL pError AS Afx_ADOError PTR, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WillChangeRecordset (BYVAL adReason AS EventReasonEnum, BYVAL adStatus AS EventStatusEnum PTR, _
           BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RecordsetChangeComplete (BYVAL adReason AS EventReasonEnum, BYVAL pError AS Afx_ADOError PTR, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WillMove (BYVAL adReason AS EventReasonEnum, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveComplete (BYVAL adReason AS EventReasonEnum, BYVAL pError AS Afx_ADOError PTR, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION EndOfRecordset (BYVAL fMoreData AS VARIANT_BOOL PTR, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION FetchProgress (BYVAL Progress AS LONG, BYVAL MaxProgress AS LONG, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION FetchComplete (BYVAL pError AS Afx_ADOError PTR, _
           BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
END TYPE
#endif

END NAMESPACE
