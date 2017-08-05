' ########################################################################################
' Microsoft Windows
' Contents: Custom error codes specific to SAPI5
' Copyright (c) 2017 José Roca
' Portions Copyright (c) Microsoft Corporation.
' All Rights Reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "win/winerror.bi"

CONST FACILITY_SAPI   = FACILITY_ITF
CONST SAPI_ERROR_BASE = &h5000

#define MAKE_SAPI_HRESULT(sev, err)    MAKE_HRESULT(sev, FACILITY_SAPI, err)
#define MAKE_SAPI_ERROR(err)           MAKE_SAPI_HRESULT(SEVERITY_ERROR, err + SAPI_ERROR_BASE)
#define MAKE_SAPI_SCODE(scode)         MAKE_SAPI_HRESULT(SEVERITY_SUCCESS, scode + SAPI_ERROR_BASE)

'/*** SPERR_UNINITIALIZED                                   &h80045001    -2147201023
'*   The object has not been properly initialized.
'*/
#define SPERR_UNINITIALIZED                                MAKE_SAPI_ERROR(&h001)

'/*** SPERR_ALREADY_INITIALIZED                             &h80045002    -2147201022
'*   The object has already been initialized.
'*/
#define SPERR_ALREADY_INITIALIZED                          MAKE_SAPI_ERROR(&h002)

'/*** SPERR_UNSUPPORTED_FORMAT                              &h80045003    -2147201021
'*   The caller has specified an unsupported format.
'*/
#define SPERR_UNSUPPORTED_FORMAT                           MAKE_SAPI_ERROR(&h003)

'/*** SPERR_INVALID_FLAGS                                   &h80045004    -2147201020
'*   The caller has specified invalid flags for this operation.
'*/
#define SPERR_INVALID_FLAGS                                MAKE_SAPI_ERROR(&h004)

'/*** SP_END_OF_STREAM                                      &h00045005    282629
'*   The operation has reached the end of stream.
'*/
#define SP_END_OF_STREAM                                   MAKE_SAPI_SCODE(&h005)

'/*** SPERR_DEVICE_BUSY                                     &h80045006    -2147201018
'*   The wave device is busy.
'*/
#define SPERR_DEVICE_BUSY                                  MAKE_SAPI_ERROR(&h006)

'/*** SPERR_DEVICE_NOT_SUPPORTED                            &h80045007    -2147201017
'*   The wave device is not supported.
'*/
#define SPERR_DEVICE_NOT_SUPPORTED                         MAKE_SAPI_ERROR(&h007)

'/*** SPERR_DEVICE_NOT_ENABLED                              &h80045008    -2147201016
'*   The wave device is not enabled.
'*/
#define SPERR_DEVICE_NOT_ENABLED                           MAKE_SAPI_ERROR(&h008)

'/*** SPERR_NO_DRIVER                                       &h80045009    -2147201015
'*   There is no wave driver installed.
'*/
#define SPERR_NO_DRIVER                                    MAKE_SAPI_ERROR(&h009)

'/*** SPERR_FILEMUSTBEUNICODE                               &h8004500a    -2147201014
'*   The file must be Unicode.
'*/
#define SPERR_FILE_MUST_BE_UNICODE                         MAKE_SAPI_ERROR(&h00a)

'/*** SP_INSUFFICIENTDATA                                   &h0004500b    282635
'*
'*/
#define SP_INSUFFICIENT_DATA                               MAKE_SAPI_SCODE(&h00b)

'/*** SPERR_INVALID_PHRASE_ID                               &h8004500c    -2147201012
'*   The phrase ID specified does not exist or is out of range.
'*/
#define SPERR_INVALID_PHRASE_ID                            MAKE_SAPI_ERROR(&h00c)

'/*** SPERR_BUFFER_TOO_SMALL                                &h8004500d    -2147201011
'*   The caller provided a buffer too small to return a result.
'*/
#define SPERR_BUFFER_TOO_SMALL                             MAKE_SAPI_ERROR(&h00d)

'/*** SPERR_FORMAT_NOT_SPECIFIED                            &h8004500e    -2147201010
'*   Caller did not specify a format prior to opening a stream.
'*/
#define SPERR_FORMAT_NOT_SPECIFIED                         MAKE_SAPI_ERROR(&h00e)

'/*** SPERR_AUDIO_STOPPED                                   &h8004500f    -2147201009
'*   This method is deprecated. Use SP_AUDIO_STOPPED instead.
'*/
#define SPERR_AUDIO_STOPPED                                MAKE_SAPI_ERROR(&h00f)

'/*** SP_AUDIO_PAUSED                                       &h00045010    282640
'*   This will be returned only on input (read) streams when the stream is paused.  Reads on
'*   paused streams will not block, and this return code indicates that all of the data has been
'*   removed from the stream.
'*/
#define SP_AUDIO_PAUSED                                    MAKE_SAPI_SCODE(&h010)

'/*** SPERR_RULE_NOT_FOUND                                  &h80045011    -2147201007
'*   Invalid rule name passed to ActivateGrammar.
'*/
#define SPERR_RULE_NOT_FOUND                               MAKE_SAPI_ERROR(&h011)

'/*** SPERR_TTS_ENGINE_EXCEPTION                            &h80045012    -2147201006
'*   An exception was raised during a call to the current TTS driver.
'*/
#define SPERR_TTS_ENGINE_EXCEPTION                         MAKE_SAPI_ERROR(&h012)

'/*** SPERR_TTS_NLP_EXCEPTION                               &h80045013    -2147201005
'*   An exception was raised during a call to an application sentence filter.
'*/
#define SPERR_TTS_NLP_EXCEPTION                            MAKE_SAPI_ERROR(&h013)

'/*** SPERR_ENGINE_BUSY                                     &h80045014    -2147201004
'*   In speech recognition, the current method can not be performed while
'*   a grammar rule is active.
'*/
#define SPERR_ENGINE_BUSY                                  MAKE_SAPI_ERROR(&h014)

'/*** SP_AUDIO_CONVERSION_ENABLED                           &h00045015    282645
'*   The operation was successful, but only with automatic stream format conversion.
'*/
#define SP_AUDIO_CONVERSION_ENABLED                        MAKE_SAPI_SCODE(&h015)

'/*** SP_NO_HYPOTHESIS_AVAILABLE                            &h00045016    282646
'*   There is currently no hypothesis recognition available.
'*/
#define SP_NO_HYPOTHESIS_AVAILABLE                         MAKE_SAPI_SCODE(&h016)

'/*** SPERR_CANT_CREATE                                     &h80045017    -2147201001
'*   Can not create a new object instance for the specified object category.
'*/
#define SPERR_CANT_CREATE                                  MAKE_SAPI_ERROR(&h017)

'/*** SP_ALREADY_IN_LEX                                     &h00045018    282648
'*   The word, pronunciation, or POS pair being added is already in lexicon.
'*/
#define SP_ALREADY_IN_LEX                                  MAKE_SAPI_SCODE(&h018)

'/*** SPERR_NOT_IN_LEX                                      &h80045019    -2147200999
'*   The word does not exist in the lexicon.
'*/
#define SPERR_NOT_IN_LEX                                   MAKE_SAPI_ERROR(&h019)

'/*** SP_LEX_NOTHING_TO_SYNC                                &h0004501a    282650
'*   The client is currently synced with the lexicon.
'*/
#define SP_LEX_NOTHING_TO_SYNC                             MAKE_SAPI_SCODE(&h01a)

'/*** SPERR_LEX_VERY_OUT_OF_SYNC                            &h8004501b    -2147200997
'*   The client is excessively out of sync with the lexicon. Mismatches may not be incrementally sync'd.
'*/
#define SPERR_LEX_VERY_OUT_OF_SYNC                         MAKE_SAPI_ERROR(&h01b)

'/*** SPERR_UNDEFINED_FORWARD_RULE_REF                      &h8004501c    -2147200996
'*   A rule reference in a grammar was made to a named rule that was never defined.
'*/
#define SPERR_UNDEFINED_FORWARD_RULE_REF                   MAKE_SAPI_ERROR(&h01c)

'/*** SPERR_EMPTY_RULE                                      &h8004501d    -2147200995
'*   A non-dynamic grammar rule that has no body.
'*/
#define SPERR_EMPTY_RULE                                   MAKE_SAPI_ERROR(&h01d)

'/*** SPERR_GRAMMAR_COMPILER_INTERNAL_ERROR                 &h8004501e    -2147200994
'*   The grammar compiler failed due to an internal state error.
'*/
#define SPERR_GRAMMAR_COMPILER_INTERNAL_ERROR              MAKE_SAPI_ERROR(&h01e)

'/*** SPERR_RULE_NOT_DYNAMIC                                &h8004501f    -2147200993
'*   An attempt was made to modify a non-dynamic rule.
'*/
#define SPERR_RULE_NOT_DYNAMIC                             MAKE_SAPI_ERROR(&h01f)

'/*** SPERR_DUPLICATE_RULE_NAME                             &h80045020    -2147200992
'*   A rule name was duplicated.
'*/
#define SPERR_DUPLICATE_RULE_NAME                          MAKE_SAPI_ERROR(&h020)

'/*** SPERR_DUPLICATE_RESOURCE_NAME                         &h80045021    -2147200991
'*   A resource name was duplicated for a given rule.
'*/
#define SPERR_DUPLICATE_RESOURCE_NAME                      MAKE_SAPI_ERROR(&h021)

'/*** SPERR_TOO_MANY_GRAMMARS                               &h80045022    -2147200990
'*   Too many grammars have been loaded.
'*/
#define SPERR_TOO_MANY_GRAMMARS                            MAKE_SAPI_ERROR(&h022)

'/*** SPERR_CIRCULAR_REFERENCE                              &h80045023    -2147200989
'*   Circular reference in import rules of grammars.
'*/
#define SPERR_CIRCULAR_REFERENCE                           MAKE_SAPI_ERROR(&h023)

'/*** SPERR_INVALID_IMPORT                                  &h80045024    -2147200988
'*   A rule reference to an imported grammar that could not be resolved.
'*/
#define SPERR_INVALID_IMPORT                               MAKE_SAPI_ERROR(&h024)

'/*** SPERR_INVALID_WAV_FILE                                &h80045025    -2147200987
'*   The format of the WAV file is not supported.
'*/
#define SPERR_INVALID_WAV_FILE                             MAKE_SAPI_ERROR(&h025)

'/*** SP_REQUEST_PENDING                                    &h00045026    282662
'*   This success code indicates that an SR method called with the SPRIF_ASYNC flag is
'*   being processed.  When it has finished processing, an SPFEI_ASYNC_COMPLETED event will be generated.
'*/
#define SP_REQUEST_PENDING                                 MAKE_SAPI_SCODE(&h026)

'/*** SPERR_ALL_WORDS_OPTIONAL                              &h80045027    -2147200985
'*   A grammar rule was defined with a null path through the rule.  That is, it is possible
'*   to satisfy the rule conditions with no words.
'*/
#define SPERR_ALL_WORDS_OPTIONAL                           MAKE_SAPI_ERROR(&h027)

'/*** SPERR_INSTANCE_CHANGE_INVALID                         &h80045028    -2147200984
'*   It is not possible to change the current engine or input.  This occurs in the
'*   following cases:
'*
'*       1) SelectEngine called while a recognition context exists, or
'*       2) SetInput called in the shared instance case.
'*/
#define SPERR_INSTANCE_CHANGE_INVALID                      MAKE_SAPI_ERROR(&h028)

'/*** SPERR_RULE_NAME_ID_CONFLICT                          &h80045029    -2147200983
'*   A rule exists with matching IDs (names) but different names (IDs).
'*/
#define SPERR_RULE_NAME_ID_CONFLICT                        MAKE_SAPI_ERROR(&h029)

'/*** SPERR_NO_RULES                                       &h8004502a    -2147200982
'*   A grammar contains no top-level, dynamic, or exported rules.  There is no possible
'*   way to activate or otherwise use any rule in this grammar.
'*/
#define SPERR_NO_RULES                                     MAKE_SAPI_ERROR(&h02a)

'/*** SPERR_CIRCULAR_RULE_REF                              &h8004502b    -2147200981
'*   Rule 'A' refers to a second rule 'B' which, in turn, refers to rule 'A'.
'*/
#define SPERR_CIRCULAR_RULE_REF                            MAKE_SAPI_ERROR(&h02b)

'/*** SP_NO_PARSE_FOUND                                    &h0004502c    282668
'*   Parse path cannot be parsed given the currently active rules.
'*/
#define SP_NO_PARSE_FOUND                                  MAKE_SAPI_SCODE(&h02c)

'/*** SPERR_NO_PARSE_FOUND                                 &h8004502d    -2147200979
'*   Parse path cannot be parsed given the currently active rules.
'*/
#define SPERR_INVALID_HANDLE                               MAKE_SAPI_ERROR(&h02d)

'/*** SPERR_REMOTE_CALL_TIMED_OUT                          &h8004502e    -2147200978
'*   A marshaled remote call failed to respond.
'*/
#define SPERR_REMOTE_CALL_TIMED_OUT                        MAKE_SAPI_ERROR(&h02e)

'/*** SPERR_AUDIO_BUFFER_OVERFLOW                           &h8004502f    -2147200977
'*   This will only be returned on input (read) streams when the stream is paused because
'*   the SR driver has not retrieved data recently.
'*/
#define SPERR_AUDIO_BUFFER_OVERFLOW                        MAKE_SAPI_ERROR(&h02f)

'/*** SPERR_NO_AUDIO_DATA                                   &h80045030    -2147200976
'*   The result does not contain any audio, nor does the portion of the element chain of the result
'*   contain any audio.
'*/
#define SPERR_NO_AUDIO_DATA                                MAKE_SAPI_ERROR(&h030)

'/*** SPERR_DEAD_ALTERNATE                                  &h80045031    -2147200975
'*   This alternate is no longer a valid alternate to the result it was obtained from.
'*   Returned from ISpPhraseAlt methods.
'*/
#define SPERR_DEAD_ALTERNATE                               MAKE_SAPI_ERROR(&h031)

'/*** SPERR_HIGH_LOW_CONFIDENCE                             &h80045032    -2147200974
'*   The result does not contain any audio, nor does the portion of the element chain of the result
'*   contain any audio.  Returned from ISpResult::GetAudio and ISpResult::SpeakAudio.
'*/
#define SPERR_HIGH_LOW_CONFIDENCE                          MAKE_SAPI_ERROR(&h032)

'/*** SPERR_INVALID_FORMAT_STRING                           &h80045033    -2147200973
'*   The XML format string for this RULEREF is invalid, e.g. not a GUID or REFCLSID.
'*/
#define SPERR_INVALID_FORMAT_STRING                        MAKE_SAPI_ERROR(&h033)

'/*** SP_UNSUPPORTED_ON_STREAM_INPUT                        &h00045034    282676
'*   The operation is not supported for stream input.
'*/
#define SP_UNSUPPORTED_ON_STREAM_INPUT                     MAKE_SAPI_SCODE(&h034)

'/*** SPERR_APPLEX_READ_ONLY                                &h80045035    -2147200971
'*   The operation is invalid for all but newly created application lexicons.
'*/
#define SPERR_APPLEX_READ_ONLY                             MAKE_SAPI_ERROR(&h035)

'/*** SPERR_NO_TERMINATING_RULE_PATH                        &h80045036    -2147200970
'*
'*/
#define SPERR_NO_TERMINATING_RULE_PATH                     MAKE_SAPI_ERROR(&h036)

'/*** SP_WORD_EXISTS_WITHOUT_PRONUNCIATION                  &h00045037    282679
'*   The word exists but without pronunciation.
'*/
#define SP_WORD_EXISTS_WITHOUT_PRONUNCIATION               MAKE_SAPI_SCODE(&h037)

'/*** SPERR_STREAM_CLOSED                                   &h80045038    -2147200968
'*   An operation was attempted on a stream object that has been closed.
'*/
#define SPERR_STREAM_CLOSED                                MAKE_SAPI_ERROR(&h038)

'// --- The following error codes are taken directly from WIN32  ---

'/*** SPERR_NO_MORE_ITEMS                                   &h80045039    -2147200967
'*   When enumerating items, the requested index is greater than the count of items.
'*/
#define SPERR_NO_MORE_ITEMS                                MAKE_SAPI_ERROR(&h039)

'/*** SPERR_NOT_FOUND                                       &h8004503a    -2147200966
'*   The requested data item (data key, value, etc.) was not found.
'*/
#define SPERR_NOT_FOUND                                    MAKE_SAPI_ERROR(&h03a)

'/*** SPERR_INVALID_AUDIO_STATE                             &h8004503b    -2147200965
'*   Audio state passed to SetState() is invalid.
'*/
#define SPERR_INVALID_AUDIO_STATE                          MAKE_SAPI_ERROR(&h03b)

'/*** SPERR_GENERIC_MMSYS_ERROR                             &h8004503c    -2147200964
'*   A generic MMSYS error not caught by _MMRESULT_TO_HRESULT.
'*/
#define SPERR_GENERIC_MMSYS_ERROR                          MAKE_SAPI_ERROR(&h03c)

'/*** SPERR_MARSHALER_EXCEPTION                             &h8004503d    -2147200963
'*   An exception was raised during a call to the marshaling code.
'*/
#define SPERR_MARSHALER_EXCEPTION                          MAKE_SAPI_ERROR(&h03d)

'/*** SPERR_NOT_DYNAMIC_GRAMMAR                             &h8004503e    -2147200962
'*   Attempt was made to manipulate a non-dynamic grammar.
'*/
#define SPERR_NOT_DYNAMIC_GRAMMAR                          MAKE_SAPI_ERROR(&h03e)

'/*** SPERR_AMBIGUOUS_PROPERTY                              &h8004503f    -2147200961
'*   Cannot add ambiguous property.
'*/
#define SPERR_AMBIGUOUS_PROPERTY                           MAKE_SAPI_ERROR(&h03f)

'/*** SPERR_INVALID_REGISTRY_KEY                            &h80045040    -2147200960
'*   The key specified is invalid.
'*/
#define SPERR_INVALID_REGISTRY_KEY                         MAKE_SAPI_ERROR(&h040)

'/*** SPERR_INVALID_TOKEN_ID                                &h80045041    -2147200959
'*   The token specified is invalid.
'*/
#define SPERR_INVALID_TOKEN_ID                             MAKE_SAPI_ERROR(&h041)

'/*** SPERR_XML_BAD_SYNTAX                                  &h80045042    -2147200958
'*   The xml parser failed due to bad syntax.
'*/
#define SPERR_XML_BAD_SYNTAX                               MAKE_SAPI_ERROR(&h042)

'/*** SPERR_XML_RESOURCE_NOT_FOUND                          &h80045043    -2147200957
'*   The xml parser failed to load a required resource (e.g., voice, phoneconverter, etc.).
'*/
#define SPERR_XML_RESOURCE_NOT_FOUND                       MAKE_SAPI_ERROR(&h043)

'/*** SPERR_TOKEN_IN_USE                                    &h80045044    -2147200956
'*   Attempted to remove registry data from a token that is already in use elsewhere.
'*/
#define SPERR_TOKEN_IN_USE                                 MAKE_SAPI_ERROR(&h044)

'/*** SPERR_TOKEN_DELETED                                   &h80045045    -2147200955
'*   Attempted to perform an action on an object token that has had associated registry key deleted.
'*/
#define SPERR_TOKEN_DELETED                                MAKE_SAPI_ERROR(&h045)

'/*** SPERR_MULTI_LINGUAL_NOT_SUPPORTED                     &h80045046    -2147200954
'*   The selected voice was registered as multi-lingual. SAPI does not support multi-lingual registration.
'*/
#define SPERR_MULTI_LINGUAL_NOT_SUPPORTED                  MAKE_SAPI_ERROR(&h046)

'/*** SPERR_EXPORT_DYNAMIC_RULE                             &h80045047    -2147200953
'*   Exported rules cannot refer directly or indirectly to a dynamic rule.
'*/
#define SPERR_EXPORT_DYNAMIC_RULE                          MAKE_SAPI_ERROR(&h047)

'/*** SPERR_STGF_ERROR                                      &h80045048    -2147200952
'*   Error parsing the SAPI Text Grammar Format (XML grammar).
'*/
#define SPERR_STGF_ERROR                                   MAKE_SAPI_ERROR(&h048)

'/*** SPERR_WORDFORMAT_ERROR                                &h80045049    -2147200951
'*   Incorrect word format, probably due to incorrect pronunciation string.
'*/
#define SPERR_WORDFORMAT_ERROR                             MAKE_SAPI_ERROR(&h049)

'/*** SPERR_STREAM_NOT_ACTIVE                               &h8004504a    -2147200950
'*   Methods associated with active audio stream cannot be called unless stream is active.
'*/
#define SPERR_STREAM_NOT_ACTIVE                            MAKE_SAPI_ERROR(&h04a)

'/*** SPERR_ENGINE_RESPONSE_INVALID                         &h8004504b    -2147200949
'*   Arguments or data supplied by the engine are in an invalid format or are inconsistent.
'*/
#define SPERR_ENGINE_RESPONSE_INVALID                      MAKE_SAPI_ERROR(&h04b)

'/*** SPERR_SR_ENGINE_EXCEPTION                             &h8004504c    -2147200948
'*   An exception was raised during a call to the current SR engine.
'*/
#define SPERR_SR_ENGINE_EXCEPTION                          MAKE_SAPI_ERROR(&h04c)

'/*** SPERR_STREAM_POS_INVALID                              &h8004504d    -2147200947
'*   Stream position information supplied from engine is inconsistent.
'*/
#define SPERR_STREAM_POS_INVALID                           MAKE_SAPI_ERROR(&h04d)

'/*** SP_RECOGNIZER_INACTIVE                                &h0004504e    282702
'*   Operation could not be completed because the recognizer is inactive. It is inactive either
'*   because the recognition state is currently inactive or because no rules are active .
'*/
#define SP_RECOGNIZER_INACTIVE                             MAKE_SAPI_SCODE(&h04e)

'/*** SPERR_REMOTE_CALL_ON_WRONG_THREAD                     &h8004504f    -2147200945
'*   When making a remote call to the server, the call was made on the wrong thread.
'*/
#define SPERR_REMOTE_CALL_ON_WRONG_THREAD                  MAKE_SAPI_ERROR(&h04f)

'/*** SPERR_REMOTE_PROCESS_TERMINATED                       &h80045050    -2147200944
'*   The remote process terminated unexpectedly.
'*/
#define SPERR_REMOTE_PROCESS_TERMINATED                    MAKE_SAPI_ERROR(&h050)

'/*** SPERR_REMOTE_PROCESS_ALREADY_RUNNING                  &h80045051    -2147200943
'*   The remote process is already running; it cannot be started a second time.
'*/
#define SPERR_REMOTE_PROCESS_ALREADY_RUNNING               MAKE_SAPI_ERROR(&h051)

'/*** SPERR_LANGID_MISMATCH                                 &h80045052    -2147200942
'*   An attempt to load a CFG grammar with a LANGID different than other loaded grammars.
'*/
#define SPERR_LANGID_MISMATCH                              MAKE_SAPI_ERROR(&h052)

'/*** SP_PARTIAL_PARSE_FOUND                               &h00045053    282707
'*   A grammar-ending parse has been found that does not use all available words.
'*/
#define SP_PARTIAL_PARSE_FOUND                             MAKE_SAPI_SCODE(&h053)

'/*** SPERR_NOT_TOPLEVEL_RULE                              &h80045054    -2147200940
'*   An attempt to deactivate or activate a non-toplevel rule.
'*/
#define SPERR_NOT_TOPLEVEL_RULE                            MAKE_SAPI_ERROR(&h054)

'/*** SP_NO_RULE_ACTIVE                                    &h00045055    282709
'*   An attempt to parse when no rule was active.
'*/
#define SP_NO_RULE_ACTIVE                                  MAKE_SAPI_SCODE(&h055)

'/*** SPERR_LEX_REQUIRES_COOKIE                            &h80045056    -2147200938
'*   An attempt to ask a container lexicon for all words at once.
'*/
#define SPERR_LEX_REQUIRES_COOKIE                          MAKE_SAPI_ERROR(&h056)

'/*** SP_STREAM_UNINITIALIZED                              &h00045057    282711
'*   An attempt to activate a rule/dictation/etc without calling SetInput
'*   first in the inproc case.
'*/
#define SP_STREAM_UNINITIALIZED                            MAKE_SAPI_SCODE(&h057)


'// Error x058 is not used in SAPI 5.0


'/*** SPERR_UNSUPPORTED_LANG                               &h80045059    -2147200935
'*   The requested language is not supported.
'*/
#define SPERR_UNSUPPORTED_LANG                             MAKE_SAPI_ERROR(&h059)

'/*** SPERR_VOICE_PAUSED                                   &h8004505a    -2147200934
'*   The operation cannot be performed because the voice is currently paused.
'*/
#define SPERR_VOICE_PAUSED                                 MAKE_SAPI_ERROR(&h05a)

'/*** SPERR_AUDIO_BUFFER_UNDERFLOW                          &h8004505b    -2147200933
'*   This will only be returned on input (read) streams when the real time audio device
'*   stops returning data for a long period of time.
'*/
#define SPERR_AUDIO_BUFFER_UNDERFLOW                       MAKE_SAPI_ERROR(&h05b)

'/*** SPERR_AUDIO_STOPPED_UNEXPECTEDLY                     &h8004505c    -2147200932
'*   An audio device stopped returning data from the Read() method even though it was in
'*   the run state.  This error is only returned in the END_SR_STREAM event.
'*/
#define SPERR_AUDIO_STOPPED_UNEXPECTEDLY                   MAKE_SAPI_ERROR(&h05c)

'/*** SPERR_NO_WORD_PRONUNCIATION                           &h8004505d    -2147200931
'*   The SR engine is unable to add this word to a grammar. The application may need to supply
'*   an explicit pronunciation for this word.
'*/
#define SPERR_NO_WORD_PRONUNCIATION                        MAKE_SAPI_ERROR(&h05d)

'/*** SPERR_ALTERNATES_WOULD_BE_INCONSISTENT                &h8004505e    -2147200930
'*   An attempt to call ScaleAudio on a recognition result having previously
'*   called GetAlternates. Allowing the call to succeed would result in
'*   the previously created alternates located in incorrect audio stream positions.
'*/
#define SPERR_ALTERNATES_WOULD_BE_INCONSISTENT             MAKE_SAPI_ERROR(&h05e)

'/*** SPERR_NOT_SUPPORTED_FOR_SHARED_RECOGNIZER             &h8004505f    -2147200929
'*   The method called is not supported for the shared recognizer.
'*   For example, ISpRecognizer::GetInputStream().
'*/
#define SPERR_NOT_SUPPORTED_FOR_SHARED_RECOGNIZER          MAKE_SAPI_ERROR(&h05f)

'/*** SPERR_TIMEOUT                                         &h80045060    -2147200928
'*   A task could not complete because the SR engine had timed out.
'*/
#define SPERR_TIMEOUT                                      MAKE_SAPI_ERROR(&h060)

'/*** SPERR_REENTER_SYNCHRONIZE                             &h80045061    -2147200927
'*   A SR engine called synchronize while inside of a synchronize call.
'*/
#define SPERR_REENTER_SYNCHRONIZE                          MAKE_SAPI_ERROR(&h061)

'/*** SPERR_STATE_WITH_NO_ARCS                              &h80045062    -2147200926
'*   The grammar contains a node no arcs.
'*/
#define SPERR_STATE_WITH_NO_ARCS                           MAKE_SAPI_ERROR(&h062)

'/*** SPERR_NOT_ACTIVE_SESSION                              &h80045063    -2147200925
'*   Neither audio output and input is supported for non-active console sessions.
'*/
#define SPERR_NOT_ACTIVE_SESSION                           MAKE_SAPI_ERROR(&h063)

'/*** SPERR_ALREADY_DELETED                                 &h80045064    -2147200924
'*   The object is a stale reference and is invalid to use.
'*   For example having a ISpeechGrammarRule object reference and then calling
'*   ISpeechRecoGrammar::Reset() will cause the rule object to be invalidated.
'*   Calling any methods after this will result in this error.
'*/
#define SPERR_ALREADY_DELETED                              MAKE_SAPI_ERROR(&h064)

'/*** SP_AUDIO_STOPPED                                      &h00045065    282725
'*   This can be returned from Read or Write calls audio streams when the stream is stopped.
'*/
#define SP_AUDIO_STOPPED                                   MAKE_SAPI_SCODE(&h065)

'/*** SPERR_RECOXML_GENERATION_FAIL                             &h80045066    -2147200922
'*   The Recognition Parse Tree couldn't be genrated.
'*   For example, that the rule name begins with a digit.
'*   XML parser doesn't allow element name beginning with a digit.
'*/
#define SPERR_RECOXML_GENERATION_FAIL                      MAKE_SAPI_ERROR(&h066)

'/*** SPERR_SML_GENERATION_FAIL                             &h80045067    -2147200921
'*   The SML couldn't be genrated.
'*   For example, the transformation xslt template is not well formed.
'*/
#define SPERR_SML_GENERATION_FAIL                          MAKE_SAPI_ERROR(&h067)

'/*** SPERR_NOT_PROMPT_VOICE                                &h80045068   -2147200920
'*   The current voice is not a prompt voice, so the ISpPromptVoice
'*   functions don't work.
'*/
#define SPERR_NOT_PROMPT_VOICE                             MAKE_SAPI_ERROR(&h068)

'/*** SPERR_ROOTRULE_ALREADY_DEFINED                        &h80045069   -2147200919
'*   There is already a root rule for this grammar
'*   Defining another root rule will fail.
'*/
#define SPERR_ROOTRULE_ALREADY_DEFINED                     MAKE_SAPI_ERROR(&h069)

'/*** SPERR_SCRIPT_DISALLOWED                               &h80045070   -2147200912
'*   Support for embedded script not supported because browser security settings have disabled it
'*/
#define SPERR_SCRIPT_DISALLOWED                            MAKE_SAPI_ERROR(&h070)

'/*** SPERR_REMOTE_CALL_TIMED_OUT_START                     &h80045071    -2147200911
'*   A time out occurred starting the sapi server
'*/
#define SPERR_REMOTE_CALL_TIMED_OUT_START                  MAKE_SAPI_ERROR(&h071)

'/*** SPERR_REMOTE_CALL_TIMED_OUT_CONNECT                   &h80045072    -2147200910
'*   A timeout occurred obtaining the lock for starting or connecting to sapi server
'*/
#define SPERR_REMOTE_CALL_TIMED_OUT_CONNECT                MAKE_SAPI_ERROR(&h072)

'/*** SPERR_SECMGR_CHANGE_NOT_ALLOWED                       &h80045073    -2147200909
'*   When there is a cfg grammar loaded, we don't allow changing the security manager
'*/
#define SPERR_SECMGR_CHANGE_NOT_ALLOWED                    MAKE_SAPI_ERROR(&h073)

'/*** SP_COMPLETE_BUT_EXTENDABLE                            &h00045074    282740
'*   Parse is valid but could be extendable (internal use only)
'*/
#define SP_COMPLETE_BUT_EXTENDABLE                         MAKE_SAPI_SCODE(&h074)

'/*** SPERR_FAILED_TO_DELETE_FILE                           &h80045075    -2147200907
'*   Tried and failed to delete an existing file.
'*/
#define SPERR_FAILED_TO_DELETE_FILE                        MAKE_SAPI_ERROR(&h075)

'/*** SPERR_SHARED_ENGINE_DISABLED                          &h80045076    -2147200906
'*   The user has chosen to disable speech from running on the machine, or the
'*   system is not set up to run speech {e.g. initial setup and tutorial has not been run}.
'*/
#define SPERR_SHARED_ENGINE_DISABLED                       MAKE_SAPI_ERROR(&h076)

'/*** SPERR_RECOGNIZER_NOT_FOUND                            &h80045077    -2147200905
'*   No recognizer is installed.
'*/
#define SPERR_RECOGNIZER_NOT_FOUND                         MAKE_SAPI_ERROR(&h077)

'/*** SPERR_AUDIO_NOT_FOUND                                 &h80045078    -2147200904
'*   No audio device is installed.
'*/
#define SPERR_AUDIO_NOT_FOUND                              MAKE_SAPI_ERROR(&h078)

'/*** SPERR_NO_VOWEL                                        &h80045079    -2147200903
'*   No Vowel in a word
'*/
#define SPERR_NO_VOWEL                                     MAKE_SAPI_ERROR(&h079)

'/*** SPERR_UNSUPPORTED_PHONEME                             &h8004507A    -2147200902
'*   Unknown phoneme
'*/
#define SPERR_UNSUPPORTED_PHONEME                          MAKE_SAPI_ERROR(&h07A)

'/*** SP_NO_RULES_TO_ACTIVATE                               &h0004507B    282747
'*   The grammar does not have any root or top-level active rules to activate.
'*/
#define SP_NO_RULES_TO_ACTIVATE                            MAKE_SAPI_SCODE(&h07B)

'/*** SP_NO_WORDENTRY_NOTIFICATION                          &h0004507C    282748
'*   The engine does not need SAPI word entry handles for this grammar
'*/
#define SP_NO_WORDENTRY_NOTIFICATION                       MAKE_SAPI_SCODE(&h07C)


'/*** SPERR_WORD_NEEDS_NORMALIZATION                        &h8004507D    -2147200899
'*   The word passed to the GetPronunciations interface needs normalizing first
'*/
#define SPERR_WORD_NEEDS_NORMALIZATION                    MAKE_SAPI_ERROR(&h07D)

'/*** SPERR_CANNOT_NORMALIZE                                &h8004507E    -2147200898
'*   The word passed to the normalize interface cannot be normalized
'*/
#define SPERR_CANNOT_NORMALIZE                           MAKE_SAPI_ERROR(&h07E)

'/*** S_LIMIT_REACHED                     &h8004507F    -2147200897
'*   The word being normalized has generated more than the maximum number of allowed normalized results
'*   Indicates that returned list is not exhaustive, but contains as many alternatives as the engine is willing to provide.
'*/
#define S_LIMIT_REACHED                                    MAKE_SAPI_SCODE(&h07F)

'/*** S_NOTSUPPORTED                     &h80045080    -2147200896
'*   We currently don't support this combination of function call + input
'*/
#define S_NOTSUPPORTED                                    MAKE_SAPI_SCODE(&h080)

'/*** SPERR_TOPIC_NOT_ADAPTABLE                            &h80045081    -2147200895
'*   This topic is not adaptable
'*/
#define SPERR_TOPIC_NOT_ADAPTABLE                         MAKE_SAPI_ERROR(&h081)

'/*** SPERR_PHONEME_CONVERSION                            &h80045082    -2147200894
'*   Cannot convert the phonemes to the specified phonetic alphabet.
'*/
#define SPERR_PHONEME_CONVERSION                         MAKE_SAPI_ERROR(&h082)

'/*** SPERR_NOT_SUPPORTED_FOR_INPROC_RECOGNIZER           &h80045083    -2147200893
'*   The method called is not supported for the in-process recognizer.
'*   For example: SetTextFeedback
'*/
#define SPERR_NOT_SUPPORTED_FOR_INPROC_RECOGNIZER         MAKE_SAPI_ERROR(&h083)

'/*** SPERR_OVERLOAD                 &h80045084         -2147200892
'*   The operation cannot be carried out due to overload and should be attempted again.
'*/
#define SPERR_OVERLOAD                              MAKE_SAPI_ERROR(&h084)

'/*** SPERR_LEX_INVALID_DATA         &h80045085         -2147200891
'*   The lexicon data is invalid or corrupted.
'*/
#define SPERR_LEX_INVALID_DATA                              MAKE_SAPI_ERROR(&h085)

'/*** SPERR_CFG_INVALID_DATA         &h80045086         -2147200890
'*   The data in the CFG grammar is invalid or corrupted
'*/
#define SPERR_CFG_INVALID_DATA                          MAKE_SAPI_ERROR(&h086)

'/*** SPERR_LEX_UNEXPECTED_FORMAT      &h80045087       -2147200889
'*   The lexicon is not in the expected format.
'*/
#define SPERR_LEX_UNEXPECTED_FORMAT                     MAKE_SAPI_ERROR(&h087)

'/*** SPERR_STRING_TOO_LONG          &h80045088         -2147200888
'*   The string is too long.
'*/
#define SPERR_STRING_TOO_LONG                           MAKE_SAPI_ERROR(&h088)

'/*** SPERR_STRING_EMPTY             &h80045089         -2147200887
'*   The string cannot be empty.
'*/
#define SPERR_STRING_EMPTY                              MAKE_SAPI_ERROR(&h089)
