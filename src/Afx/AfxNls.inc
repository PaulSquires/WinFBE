' ########################################################################################
' Microsoft Windows
' File: AfxNls.inc
' Contents: National language support functions
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "win/winnls.bi"
#INCLUDE ONCE "Afx/CWStr.inc"
USING Afx

' ========================================================================================
' Returns the abbreviated native day name for the specified day (1 = Monday, etc.).
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleAbbrevDayName (BYVAL nDay AS LONG, BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   IF nDay < 1 OR nDay > 7 THEN RETURN ""
   GetLocaleInfoW(nLocale, LOCALE_SABBREVDAYNAME1 + nDay - 1, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the abbreviated native month name for the specified month.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleAbbrevMonthName (BYVAL nMonth AS LONG, BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   IF nMonth < 1 OR nMonth > 12 THEN RETURN ""
   GetLocaleInfoW(nLocale, LOCALE_SABBREVMONTHNAME1 + nMonth - 1, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Full localized name of the country/region. Deprecated for Windows 7 and later.
' See AfxGetLocalizedCountryName.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleCountryName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SCOUNTRY, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' String used as the local monetary symbol. The maximum number of characters allowed for
' this string is 13, including a terminating null character. For example, in the United
' States, this symbol is "$".
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleCurrencySymbol (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SCURRENCY, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the native day name for the specified day (1 = Monday, etc.).
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleDayName (BYVAL nDay AS LONG, BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   IF nDay < 1 OR nDay > 7 THEN RETURN ""
   GetLocaleInfoW(nLocale, LOCALE_SDAYNAME1 + nDay - 1, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Character(s) used for the decimal separator, for example, "." in "3.14" or "," in "3,14".
' The maximum number of characters allowed for this string is four, including a terminating
' null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleDecimalSeparator (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SDECIMAL, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sizes for each group of digits to the left of the decimal. The maximum number of characters
' allowed for this string is ten, including a terminating null character. An explicit size
' is needed for each group, and sizes are separated by semicolons. If the last value is 0,
' the preceding value is repeated. For example, to group thousands, specify 3;0. Indic
' locales group the first thousand and then group by hundreds. For example, 12,34,56,789
' is represented by 3;2;0.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleDigitsGrouping (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SGROUPING, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Primary language name included in a localized display name. Deprecated in Windows 7.
' See AfxGetLocalizedDisplayName.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleDisplayName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SLANGDISPLAYNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Full English name of the country/region.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleEnglishCountryName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SENGCOUNTRY, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The full English name of the currency associated with the locale.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleEnglishCurrencyName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SENGCURRNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows 7 and later: Display name of the locale in English. Usually the display name
' consists of the language and the country/region, for example, German (Germany) for
' Deutsch (Deutschland).
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleEnglishDisplayName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SENGLISHDISPLAYNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Full English name of the language from ISO Standard 639.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleEnglishLanguageName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SENGLANGUAGE, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Three characters of the international monetary symbol specified in ISO 4217, followed by
' the character separating this string from the amount. The maximum number of characters
' allowed for this string is nine, including a terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleInternationalMonetarySymbol (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SINTLSYMBOL, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The abbreviated name of the language based entirely on the ISO Standard 639 values, in
' lowercase form, for example, "en" for English. The maximum number of characters allowed
' for this string is nine, including a terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleISO639LanguageName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SISO639LANGNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows Vista and later: Three-letter ISO language name, in lowercase form (ISO 639-2
' three-letter code for the language), for example, "eng" for English. The maximum number
' of characters allowed for this string is nine, including a terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleISO639LanguageName2 (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SISO639LANGNAME2, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Country/region name, based on ISO Standard 3166, for example, "US" for the United States.
' The maximum number of characters allowed for this string is nine, including a terminating
' null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleISO3166CountryName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SISO3166CTRYNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows Vista and later: Three-letter ISO region name (ISO 3166 three-letter code for the
' country/region), for example, "USA" for the United States. The maximum number of characters
' allowed for this string is nine, including a terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleISO3166CountryName2 (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SISO3166CTRYNAME2, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Full localized name of the language. Deprecated in Windows 7.
' See AfxGetLocalizedLanguageName.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleLanguageName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SLANGUAGE, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Character(s) used to separate list items, for example, a comma is used in many locales.
' The maximum number of characters allowed for this string is four, including a
' terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleListItemsSeparator (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SLIST, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Long date formatting string for the locale. The maximum number of characters allowed for
' this string is 80, including a terminating null character. The string can consist of a
' combination of day, month, year, and era format pictures and any string of characters
' enclosed in single quotes. Characters in single quotes remain as specified. For example,
' the Spanish (Spain) long date is "dddd, dd' de 'MMMM' de 'yyyy". Locales can define
' multiple long date formats.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleLongDateFormatString (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SLONGDATE, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Character(s) used as the monetary decimal separator. The maximum number of characters
' allowed for this string is four, including a terminating null character. For example,
' if a monetary amount is displayed as "$3.40", just as "three dollars and forty cents"
' is displayed in the United States, then the monetary decimal separator is ".".
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleMonetaryDecimalSeparator (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SMONDECIMALSEP, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sizes for each group of monetary digits to the left of the decimal. The maximum number
' of characters allowed for this string is ten, including a terminating null character.
' An explicit size is needed for each group, and sizes are separated by semicolons. If the
' last value is 0, the preceding value is repeated. For example, to group thousands,
' specify 3;0. Indic languages group the first thousand and then group by hundreds. For
' example 12,34,56,789 is represented by 3;2;0.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleMonetaryGrouping (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SMONGROUPING, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Character(s) used as the monetary separator between groups of digits to the left of the
' decimal. The maximum number of characters allowed for this string is four, including a
' terminating null character. Typically, the groups represent thousands. However, depending
' on the value specified for LOCALE_SMONGROUPING, they can represent something else.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleMonetaryThousandsSeparator (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SMONTHOUSANDSEP, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows 7 and later: Format string for displaying only the month and the day. The
' formatting is similar to that defined for LOCALE_SLONGDATE. For example, if the month/day
' pattern is the full month name followed by the day number with leading zeros, as in
' "September 03", the format string is "MMMM dd". The string can consist of a combination
' of day and month format pictures and any string of characters enclosed in single quotes.
' Characters in single quotes remain as specified, for example, 'de' for Spanish (Spain).
' A locale can specify only one month/day format.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleMonthDayFormatString (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SMONTHDAY, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Returns the native month name for the specified month.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleMonthName (BYVAL nMonth AS LONG, BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   IF nMonth < 1 OR nMonth > 12 THEN RETURN ""
   GetLocaleInfoW(nLocale, LOCALE_SMONTHNAME1 + nMonth - 1, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows Vista and later: Locale name, a multi-part tag to uniquely identify the locale.
' The maximum number of characters allowed for this string is 85, including a terminating
' null character. The tag is based on the language tagging conventions of RFC 4646.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows Vista and later: String value for "Not a number", for example, "Nan" for the
' English (United States) locale. There is no limit on the number of characters allowed
' for this string.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleNanString (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM nLen AS LONG, wstrBuf AS CWSTR
   nLen = GetLocaleInfoW(nLocale, LOCALE_SNAN, BYVAL NULL, 0)
   IF nLen = 0 THEN RETURN ""
   wstrBuf = SPACE(nLen)
   nLen = GetLocaleInfoW(nLocale, LOCALE_SNAN, @wstrBuf, nLen)
   RETURN TRIM(wstrBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Native name of the country/region, for example, España for Spain. The maximum number of
' characters allowed for this string is 80, including a terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleNativeCountryName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SNATIVECTRYNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The native name of the currency associated with the locale, in the native language of
' the locale. There is no limit on the number of characters allowed for this string.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleNativeCurrencyName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SNATIVECURRNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Native equivalents of ASCII 0 through 9. The maximum number of characters allowed for
' this string is eleven, including a terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleNativeDigits (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SNATIVEDIGITS, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows 7 and later: Display name of the locale in its native language, for example,
' Deutsch (Deutschland) for the locale German (Germany).
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleNativeDisplayName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SNATIVEDISPLAYNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Native name of the language, for example, ??????? for Armenian (Armenia). The maximum
' number of characters allowed for this string is 80, including a terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleNativeLanguageName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SNATIVELANGNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows Vista and later: String value for "negative infinity", for example, "-Infinity"
' for the English (United States) locale. There is no limit on the number of characters
' allowed for this string.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleNegativeInfinityString (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM nLen AS LONG, wstrBuf AS CWSTR
   nLen = GetLocaleInfoW(nLocale, LOCALE_SNEGINFINITY, NULL, 0)
   IF nLen = 0 THEN RETURN ""
   wstrBuf = SPACE(nLen)
   nLen = GetLocaleInfoW(nLocale, LOCALE_SNEGINFINITY, @wstrBuf, nLen)
   RETURN TRIM(wstrBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' String value for the negative sign, for example, "-" for the English (United States) locale.
' The maximum number of characters allowed for this string is five, including a terminating
' null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleNegativeSign (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SNEGATIVESIGN, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows 7 and later: OpenType language tag used to retrieve culturally appropriate
' typographic features from a font.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleOpenTypeLanguageTag (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SOPENTYPELANGUAGETAG, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows 7 and later: Symbol used to indicate the permille (U+2030) symbol, that is, ‰.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocalePermilleSymbol (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SPERMILLE, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows Vista and later: String value for "positive infinity", for example, "Infinity"
' for the English (United States) locale. There is no limit on the number of characters
' allowed for the string.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocalePositiveInfinityString (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM nLen AS LONG, wstrBuf AS CWSTR
   nLen = GetLocaleInfoW(nLocale, LOCALE_SPOSINFINITY, BYVAL NULL, 0)
   IF nLen = 0 THEN RETURN ""
   wstrBuf = SPACE(nLen)
   nLen = GetLocaleInfoW(nLocale, LOCALE_SPOSINFINITY, @wstrBuf, nLen)
   RETURN TRIM(wstrBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' String value for the positive sign, for example, "+" for the English (United States)
' locale. The maximum number of characters allowed for this string is five, including a
' terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocalePositiveSign (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SPOSITIVESIGN, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows Vista and later: A string representing a list of scripts, using the 4-character
' notation used in ISO 15924. Each script name consists of four Latin characters and the
' list is arranged in alphabetical order with each name, including the last, followed by a
' semicolon.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleScriptsList (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM nLen AS LONG, wstrBuf AS CWSTR
   nLen = GetLocaleInfoW(nLocale, LOCALE_SSCRIPTS, BYVAL NULL, 0)
   IF nLen = 0 THEN RETURN ""
   wstrBuf = SPACE(nLen)
   nLen = GetLocaleInfoW(nLocale, LOCALE_SSCRIPTS, @wstrBuf, nLen)
   RETURN TRIM(wstrBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' Short date formatting string for the locale. The maximum number of characters allowed
' for this string is 80, including a terminating null character. The string can consist of
' a combination of day, month, year, and era format pictures. For example, "M/d/yyyy"
' indicates that September 3, 2004 is written 9/3/2004.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleShortDateFormatString (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SSHORTDATE, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows Vista and later: Short native name of the specified day of the week.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleShortestDayName (BYVAL nDay AS LONG, BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   IF nDay < 1 OR nDay > 7 THEN RETURN ""
   GetLocaleInfoW(nLocale, LOCALE_SSHORTESTDAYNAME1 + nDay - 1, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows 7 and later: Short time formatting string for the locale. Patterns are typically
' derived by removing the "ss" (seconds) value from the long time format pattern.
' For example, if the long time format is "h:mm:ss tt", the short time format is most
' likely "h:mm tt". This constant can specify multiple formats in a semicolon-delimited
' list. However, the preferred short time format should be the first value listed.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleShortTimeFormatString (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SSHORTTIME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif

' ========================================================================================
' The full localized name of the sort for the specified locale identifier, dependent on
' the language of the shell. This constant is used to determine casing and sorting behavior.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleSortName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SSORTNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Characters that are used to separate groups of digits to the left of the decimal. The
' maximum number of characters allowed for this string is four, including a terminating
' null character. Typically, these groups represent thousands. However, depending on the
' value specified for LOCALE_SGROUPING, they can represent something else.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleThousandsSeparator (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_STHOUSAND, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Character(s) for the date separator. The maximum number of characters allowed for this
' string is four, including a terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleDateSeparator (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SDATE, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Time formatting strings for the locale. The maximum number of characters allowed for
' this string is 80, including a terminating null character. The string can consist of a
' combination of hour, minute, and second format pictures.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleTimeFormatString (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_STIMEFORMAT, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Character(s) for the time separator. The maximum number of characters allowed for this
' string is four, including a terminating null character.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleTimeSeparator (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_STIME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The year-month formatting string for the locale. The maximum number of characters allowed
' for this string is 80, including a terminating null character. This string shows the
' proper format for a date string that contains only the year and the month, using format
' pictures as defined in Day, Month, Year, and Era Format Pictures.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocaleYearMonthFormatString (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SYEARMONTH, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Windows 7 and later: Full localized name of the country/region, for example, Deutschland
' for Germany. The maximum number of characters allowed for this string is 80, including a
' terminating null character. Since this name is based on the localization of the product,
' it changes for each localized version.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocalizedCountryName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SLOCALIZEDCOUNTRYNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Windows 7 and later: Full localized name of the locale for the user interface language,
' for example, Deutsch (Deutschland) for German (Germany)" There is no limit on the number
' of characters allowed for this string. Since this name is based on the localization of
' the product, it changes for each localized version.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocalizedDisplayName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SLOCALIZEDDISPLAYNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Windows Vista: Full localized primary name of the user interface language included in a
' localized display name, for example, Deutsch representing German. Since this name is
' based on the localization of the product, it changes for each localized version.
' ========================================================================================
PRIVATE FUNCTION AfxGetLocalizedLanguageName (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SLOCALIZEDDISPLAYNAME, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================

#if _WIN32_WINNT = &h0602
' ========================================================================================
' Windows 7 and later: Name of the locale to use for sorting or casing behavior.
' ========================================================================================
PRIVATE FUNCTION AfxGetSortLocale (BYVAL nLocale AS LONG = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszBuf AS WSTRING * 260
   GetLocaleInfoW(nLocale, LOCALE_SSORTLOCALE, wszBuf, SIZEOF(wszBuf))
   RETURN TRIM(wszBuf)
END FUNCTION
' ========================================================================================
#endif
