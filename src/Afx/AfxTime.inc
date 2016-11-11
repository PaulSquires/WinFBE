' ########################################################################################
' Microsoft Windows
' File: AfxTime.inc
' Contents: Date and time functions.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#INCLUDE ONCE "windows.bi"
#include once "Afx/CWStr.inc"
USING Afx

' ========================================================================================
' Returns the current local time as a FILETIME structure.
' ========================================================================================
PRIVATE FUNCTION AfxLocalFileTime () AS FILETIME
   DIM ST AS SYSTEMTIME, FT AS FILETIME
   GetLocalTime @ST
   SystemTimeToFileTime @ST, @FT
   RETURN FT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current local time as a SYSTEMTIME structure.
' ========================================================================================
PRIVATE FUNCTION AfxLocalSystemTime () AS SYSTEMTIME
   DIM ST AS SYSTEMTIME
   GetLocalTime @ST
   RETURN ST
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current local year. The valid values are 1601 through 30827.
' ========================================================================================
PRIVATE FUNCTION AfxLocalYear () AS WORD
   DIM ST AS SYSTEMTIME
   GetLocalTime @ST
   RETURN ST.wYear
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current local month. The valid values are 1 thorugh 12 (1 = January, etc.).
' ========================================================================================
PRIVATE FUNCTION AfxLocalMonth () AS WORD
   DIM ST AS SYSTEMTIME
   GetLocalTime @ST
   RETURN ST.wMonth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current local day. The valid values are 1 through 31.
' ========================================================================================
PRIVATE FUNCTION AfxLocalDay () AS WORD
   DIM ST AS SYSTEMTIME
   GetLocalTime @ST
   RETURN ST.wDay
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current local hour. The valid values are 0 through 23.
' ========================================================================================
PRIVATE FUNCTION AfxLocalHour () AS WORD
   DIM ST AS SYSTEMTIME
   GetLocalTime @ST
   RETURN ST.wHour
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the localized name of today's month.
' ========================================================================================
PRIVATE FUNCTION AfxLocalMonthName (BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, wszDate AS WSTRING * 260
   GetSystemTime(@ST)
   GetDateFormatW(lcid, NULL, @ST, "MMMM", wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the localized short name of today's month.
' ========================================================================================
PRIVATE FUNCTION AfxLocalShortMonthName (BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, wszDate AS WSTRING * 260
   GetSystemTime(@ST)
   GetDateFormatW(lcid, NULL, @ST, "MMM", wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the localized name of today.
' ========================================================================================
PRIVATE FUNCTION AfxLocalDayName (BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, wszDate AS WSTRING * 260
   GetSystemTime(@ST)
   IF lcid = 0 THEN lcid = LOCALE_USER_DEFAULT
   GetDateFormatW(lcid, NULL, @ST, "dddd", wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the localized short name of today.
' ========================================================================================
PRIVATE FUNCTION AfxLocalDayShortName (BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, wszDate AS WSTRING * 260
   GetSystemTime(@ST)
   GetDateFormatW(lcid, NULL, @ST, "ddd", wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current day of week. It is a numeric value in the range of 0-6 representing
' Sunday through Saturday: 0 = Sunday, 1 = Monday, etc.
' ========================================================================================
PRIVATE FUNCTION AfxLocalDayOfWeek () AS WORD
   DIM ST AS SYSTEMTIME
   GetLocalTime @ST
   RETURN ST.wDayOfWeek
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the localized name of the day of the week.
' Parameters:
' - nDay: A number between 1-31.
' - nMonth: A number between 1-12.
' - nYear: A four digit year, e.g. 2011.
' Return Value:
' - The localized day of week name.
' ========================================================================================
PRIVATE FUNCTION AfxLocalDayOfWeekName (BYVAL nDay AS LONG, BYVAL nMonth AS LONG, BYVAL nYear AS LONG, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME
   DIM wszDate AS WSTRING * 260
   ST.wYear = nYear
   ST.wMonth = nMonth
   ST.wDay = nDay
   GetDateFormatW(lcid, NULL, @ST, "dddd", wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the localized short name of the day of the week.
' Parameters:
' - nDay: A number between 1-31.
' - nMonth; A number between 1-12.
' - nYear: A four digit year, e.g. 2011.
' Return Value:
' - The localized day of week short name.
' ========================================================================================
PRIVATE FUNCTION AfxLocalDayOfWeekShortName (BYVAL nDay AS LONG, BYVAL nMonth AS LONG, BYVAL nYear AS LONG, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME
   DIM wszDate AS WSTRING * 260
   ST.wYear = nYear
   ST.wMonth = nMonth
   ST.wDay = nDay
   GetDateFormatW(lcid, NULL, @ST, "ddd", wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retuns the current local date based on the specified mask, e.g. "dd-MM-yyyy".
' ========================================================================================
PRIVATE FUNCTION AfxLocalDateStr (BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszDateStr AS WSTRING * 260
   GetDateFormatW(lcid, NULL, NULL, wszMask, wszDateStr, SIZEOF(wszDateStr))
   RETURN wszDateStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retuns the current local time based on the specified mask, e.g. "hh':'mm':'ss".
' ========================================================================================
PRIVATE FUNCTION AfxLocalTimeStr (BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszDateStr AS WSTRING * 260
   GetTimeFormatW(lcid, NULL, NULL, wszMask, wszDateStr, SIZEOF(wszDateStr))
   RETURN wszDateStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the localized name of the specified month.
' ========================================================================================
PRIVATE FUNCTION AfxMonthName (BYVAL nMonth AS LONG, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, wszDate AS WSTRING * 260
   GetSystemTime(@ST)
   ST.wMonth = nMonth
   GetDateFormatW(lcid, NULL, @ST, "MMMM", wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the localized short name of the specified month.
' ========================================================================================
PRIVATE FUNCTION AfxShortMonthName (BYVAL nMonth AS LONG, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, wszDate AS WSTRING * 260
   GetSystemTime(@ST)
   ST.wMonth = nMonth
   GetDateFormatW(lcid, NULL, @ST, "MMM", wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the Astronomical Day for any given date.
' Parameters:
' - nDay: A number between 1-31.
' - nMonth; A number between 1-12.
' - nYear: A four digit year, e.g. 2011.
' Return Value:
' - The Astronomical Day.
' See: http://support.microsoft.com/kb/109451/en-us
' Note: Among other things, can be used to find the number of days between any two dates, e.g.:
' PRINT AfxAstroDay(1, 3, -12400) - AfxAstroDay(28, 2, -12400)  ' Prints 2
' PRINT AfxAstroDay(1, 3, 12000) - AfxAstroDay(28, 2, -12000) ' Prints 8765822
' PRINT AfxAstroDay(28, 2, 1902) - AfxAstroDay(1, 3, 1898)  ' Prints 1459 days
' ========================================================================================
PRIVATE FUNCTION AfxAstroDay (BYVAL nDay AS LONG, BYVAL nMonth AS LONG, BYVAL nYear AS LONG) AS LONG
   DIM y AS DOUBLE = nYear + (nMonth - 2.85) / 12
   RETURN INT(INT(INT(367 * y) - 1.75 * INT(y) + nDay) - 0.75 * INT(0.01 * y)) + 1721116
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates the day of the week, Sunday through Monday, of a given date.
' Parameters:
' - nDay: A number between 1-31.
' - nMonth; A number between 1-12.
' - nYear: A four digit year, e.g. 2011.
' Return Value:
' - A number between 0-6, where 0 = Sunday, 1 = Monday, 2 = Tuesday, 3 = Wednesday,
'   4 = Thursday, 5 = Friday, 6 = Saturday.
' ========================================================================================
PRIVATE FUNCTION AfxAstroDayOfWeek (BYVAL nDay AS LONG, BYVAL nMonth AS LONG, BYVAL nYear AS LONG) AS LONG
   RETURN AfxAstroDay(nDay, nMonth, nYear) MOD 7
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a Gregorian date to a Julian date.
' Parameters:
' - nDay: A number between 1-31.
' - nMonth: A number between 1-12.
' - nYear: A four digit year, e.g. 2011.
' Return Value:
' - The converted date.
' ========================================================================================
PRIVATE FUNCTION AfxGregorianToJulian (BYVAL nDay AS LONG, BYVAL nMonth AS LONG, BYVAL nYear AS LONG) AS LONG

   DIM DaysElapsed AS LONG
   DIM InvalidLeapYears AS LONG
   DIM ValidLeapYears AS LONG
   DIM DaysOfMonths AS LONG

   ' // January or February? 13 or 14th month of the previous year.
   IF nMonth < 3 THEN nMonth += 12 : nYear -=1
   ' // Day 1 of julian calendar is 1/1/-4712 and there are 365.25 days per year
   DaysElapsed = INT((nYear + 4712) * 365.25)
   ' // Substract invalid leap years
   InvalidLeapYears = nYear \ 100
   ' // Add valid leap years
   ValidLeapYears = nYear \ 400
   ' // Days of months
   DaysOfMonths = INT(30.6 * (nMonth - 1) + .2)

   RETURN DaysElapsed - InvalidLeapYears + ValidLeapYears + DaysOfMonths + nDay

END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates the days of difference between two dates.
' Parameters:
' - nDay1, nMonth1, nYear1 : The first date
' - nDay2, nMonth2, nYear2 : The second date
' Return Value:
' - The days of difference.
' ========================================================================================
PRIVATE FUNCTION AfxDateDiff (BYVAL nDay1 AS LONG, BYVAL nMonth1 AS LONG, BYVAL nYear1 AS LONG, _
   BYVAL nDay2 AS LONG, BYVAL nMonth2 AS LONG, BYVAL nYear2 AS LONG) AS LONG
   DIM t1 AS LONG = AfxGregorianToJulian(nDay1, nMonth1, nYear1)
   DIM t2 AS LONG = AfxGregorianToJulian(nDay2, nMonth2, nYear2)
   RETURN t2 - t1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds the specified number of days to a given date.
' Parameters:
' - nDay: A number between 1-31.
' - nMonth:s A number between 1-12.
' - nYear: A four digit year, e.g. 2011.
' - nDays: The number of days to add (or substract, if the number is negative).
' Return Value:
' - The new date in Julian format. To convert it to Gregorian, call functions such
'   AfxJulianToGregorian or AfxJulianToGregorianStr.
' ========================================================================================
PRIVATE FUNCTION AfxDateAddDays (BYVAL nDay AS LONG, BYVAL nMonth AS LONG, BYVAL nYear AS LONG, BYVAL nDays AS LONG) AS LONG
   RETURN AfxGregorianToJulian(nDay, nMonth, nYear) + nDays
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the first day of the first week of a year. The year must be a 4 digit year.
' Please note that the result can be one of the last days of the previous year, e.g. the
' first day of the first week of year 2002 is 31 December 2001.
' To convert the returned julian date to gregorian use AfxJulianToGregorian or AfxJulianToGreorianStr.
' ========================================================================================
PRIVATE FUNCTION AfxWeekOne (BYVAL nYear AS LONG) AS LONG
   DIM nFirstDay AS LONG = AfxGregorianToJulian(1, 1, nYear) - 1  ' Dec. 31 of prev. year
   DO
      nFirstDay += 1
   LOOP UNTIL (nFirstDay MOD 7 + 1) = 4    ' until first Thursday of the year is found
   RETURN nFirstDay - 3                ' first day of week 1 is a Monday
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the week number for a given date. The year must be a 4 digit year.
' ========================================================================================
PRIVATE FUNCTION AfxWeekNumber (BYVAL nDay AS LONG, BYVAL nMonth AS LONG, BYVAL nYear AS LONG) AS LONG
   DIM nFirstDay AS LONG = AfxWeekOne(nYear)
   DIM nFinalDay AS LONG = AfxWeekOne(nYear + 1) - 1
   DIM nToDay AS LONG = AfxGregorianToJulian(nDay, nMonth, nYear)
   SELECT CASE nToDay
      CASE IS < nFirstDay
         ' it is week 52 or 53, but which one?
         ' therefore we need week one of previous year as a starting point
         nFirstDay = AfxWeekOne(nYear - 1)
      CASE IS > nFinalDay
         ' there is only one possibility: week number 1
         RETURN 1
   END SELECT
   RETURN ((nToDay - nFirstDay) \ 7) + 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of weeks in the year, where weeks are taken to start on Monday. Will be 52 or 53.
' ========================================================================================
PRIVATE FUNCTION AfxWeeksInYear (BYVAL nYear AS LONG) AS LONG
   DIM nWeeks AS LONG = AfxWeekNumber(31, 12, nYear)
   IF nWeeks = 53 THEN RETURN 53 ELSE RETURN 52
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if a given year is a leap year or not.
' Parameters:
' - nYear: A four digit year, e.g. 2011.
' Return Value:
' - TRUE or FALSE.
' Note: A leap year is defined as all years divisible by 4, except for years divisible by
' 100 that are not also divisible by 400. Years divisible by 400 are leap years. 2000 is a
' leap year. 1900 is not a leap year.
' ========================================================================================
PRIVATE FUNCTION AfxIsLeapYear (BYVAL nYear AS LONG) AS BOOLEAN
   IF (nYear MOD 4 = 0 AND nYear MOD 100 <> 0) OR nYear MOD 400 = 0 THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of days in the specified month.
' Parameters:
' - nMonth: A two digit month (1-12).
' - nYear: A four digit year, e.g. 2011.
' Return Value:
' - The number of days.
' ========================================================================================
PRIVATE FUNCTION AfxDaysInMonth (BYVAL nMonth AS LONG, BYVAL nYear AS LONG) AS LONG
   SELECT CASE nMonth
      CASE 2 : IF AfxIsLeapYear(nYear) THEN RETURN 29 ELSE RETURN 28
      CASE 4, 6, 9, 11 : RETURN 30
      CASE 1, 3, 5, 7, 8, 10, 12 : RETURN 31
   END SELECT
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of weeks in the specified month. Will be 4 or 5.
' ========================================================================================
PRIVATE FUNCTION AfxWeeksInMonth (BYVAL nMonth AS LONG, BYVAL nYear AS LONG) AS LONG
   DIM wim AS LONG = AfxAstroDayOfWeek(1, nMonth, nYear) - 1
   IF wim = -1 THEN wim = 6
   DIM nWeeks AS LONG = (AfxDaysInMonth(nMonth, nYear) + wim) \ 7
   RETURN nWeeks
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of days in the specified year.
' Parameters:
' - nYear: A four digit year, e.g. 2011.
' Return Value:
' - The number of days. Will be 365 or 366 (for leap years).
' ========================================================================================
PRIVATE FUNCTION AfxDaysInYear (BYVAL nYear AS LONG) AS LONG
   IF AfxIsLeapYear(nYear) THEN RETURN 366 ELSE RETURN 365
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if today is the first day of the month; False, otherwise.
' ========================================================================================
PRIVATE FUNCTION IsFirstDayOfMonth () AS BOOLEAN
   DIM ST AS SYSTEMTIME
   GetLocalTime @ST
   RETURN (ST.wDay = 1)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the date is the last day of the month; False, otherwise.
' ========================================================================================
PRIVATE FUNCTION IsLastDayOfMonth () AS BOOLEAN
   DIM ST AS SYSTEMTIME
   GetLocalTime @ST
   RETURN (ST.wDay = AfxDaysInMonth(ST.wMonth, ST.wYear))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Given a Julian date, returns the day of week.
' ========================================================================================
PRIVATE FUNCTION AfxJulianDayOfWeek (BYVAL nJulian AS LONG) AS LONG
   DIM dow AS LONG = (nJulian + 1) MOD 7
   IF dow >= 0 THEN RETURN dow ELSE RETURN dow + 7
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a Julian date to a Gregorian date.
' Parameters:
' - nJulian [in] = The julian date to convert
' - nDay [out] = The day
' - nMonth [out] = The month
' - nYear [out] = The year
' ========================================================================================
PRIVATE SUB AfxJulianToGregorian (BYVAL nJulian AS LONG, BYREF wDay AS WORD, BYREF wMonth AS WORD, BYREF wYear AS WORD)
   nJulian = nJulian + 68569
   DIM nHelp AS LONG = 4 * nJulian \ 146097
   nJulian = nJulian - ((146097 * nHelp + 3) \ 4)
   DIM nTmpYear AS LONG = 4000 * (nJulian + 1) \ 1461001
   nJulian = nJulian - (1461 * nTmpYear \ 4) + 31
   DIM nTmpMonth AS LONG = 80 * nJulian \ 2447
   wDay = nJulian - (2447 * nTmpMonth \ 80)
   wMonth = nTmpMonth + 2 - (12 * (nTmpMonth \ 11))
   wYear = 100 * (nHelp - 49) + nTmpYear + (nTmpMonth \ 11)
END SUB
' ========================================================================================

' ========================================================================================
' Converts a Julian date to a Gregorian date based on the specified mask, e.g. "dd-MM-yyyy".
' Parameters:
' - bstrMask = The format string.
' - nJulian = The Julian date to convert.
' Return Value:
' - The converted date.
' ========================================================================================
PRIVATE FUNCTION AfxJulianToGregorianStr (BYREF wszMask AS WSTRING, BYVAL nJulian AS LONG, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME
   DIM wszDate AS WSTRING * 260
   AfxJulianToGregorian(nJulian, ST.wDay, ST.wMonth, ST.wYear)
   GetDateFormatW(lcid, NULL, @ST, wszMask, wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the day of the year, where Jan 1 is the first day of the year.
' ========================================================================================
PRIVATE FUNCTION AfxDayOfYear (BYVAL nDay AS LONG, BYVAL nMonth AS LONG, BYVAL nYear AS LONG) AS LONG
   RETURN AfxAstroDay(nDay, nMonth, nYear) - AfxAstroDay(1, 1, nYear) + 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of days between two given dates.
' ========================================================================================
PRIVATE FUNCTION AfxDaysBetween (BYVAL nDay1 AS LONG, BYVAL nMonth1 AS LONG, BYVAL nYear1 AS LONG, BYVAL nDay2 AS LONG, BYVAL nMonth2 AS LONG, BYVAL nYear2 AS LONG) AS LONG
   RETURN AfxAstroDay(nDay1, nMonth1, nYear1) - AfxAstroDay(nDay2, nMonth2, nYear2)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current date in long format.
' ========================================================================================
PRIVATE FUNCTION AfxLongDate (BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME
   DIM wszDate AS WSTRING * 260
   GetLocalTime @ST
   GetDateFormatW(lcid, DATE_LONGDATE, @ST, BYVAL NULL, wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current date in short format.
' ========================================================================================
PRIVATE FUNCTION AfxShortDate (BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME
   DIM wszDate AS WSTRING * 260
   GetLocalTime @ST
   GetDateFormatW(lcid, DATE_SHORTDATE, @ST, NULL, wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the curent date and time as a QUAD (8 bytes). IN FB, a QUAD is an ULONGLONG.
' ========================================================================================
PRIVATE FUNCTION AfxQuadDateTime () AS ULONGLONG
   DIM ST AS SYSTEMTIME, FT AS FILETIME
   GetLocalTime @ST
   SystemTimeToFileTime @ST, @FT
   DIM uli AS ULARGE_INTEGER
   uli.LowPart = FT.dwLowDateTime
   uli.HighPart = FT.dwHighDateTime
   RETURN uli.QuadPart
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a date stored in a QUAD into a formatted date string. For example, to get a date
' string formated like "Wed, Aug 31 2016" use the following picture string: "ddd',' MMM dd yyy".
' IN FB, a QUAD is an ULONGLONG.
' ========================================================================================
PRIVATE FUNCTION AfxQuadDateToStr (BYREF wszMask AS WSTRING, BYVAL ullTime AS ULONGLONG, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, FT AS FILETIME
   DIM uli AS ULARGE_INTEGER
   uli.QuadPart = ullTime
   FT.dwLowDateTime = uli.LowPart
   FT.dwHighDateTime = uli.HighPart
   FileTimeToSystemTime(@FT, @ST)
   DIM wszDate AS WSTRING * 260
   GetDateFormatW(lcid, NULL, @ST, wszMask, wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a time stored in a QUAD into a formatted time string. For example, to get a time
' string like "11:29:40 PM" use the following picture string: "hh':'mm':'ss tt".
' IN FB, a QUAD is an ULONGLONG.
' ========================================================================================
PRIVATE FUNCTION AfxQuadTimeToStr (BYREF wszMask AS WSTRING, BYVAL ullTime AS ULONGLONG, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, FT AS FILETIME
   DIM uli AS ULARGE_INTEGER
   uli.QuadPart = ullTime
   FT.dwLowDateTime = uli.LowPart
   FT.dwHighDateTime = uli.HighPart
   FileTimeToSystemTime(@FT, @ST)
   DIM wszDate AS WSTRING * 260
   GetTimeFormatW(lcid, NULL, @ST, wszMask, wszDate, SIZEOF(wszDate))
   RETURN wszDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether the the system is operating in the range of daylight saving time
' for the current time zone.
' ========================================================================================
PRIVATE FUNCTION AfxTimeZoneIsDaylightSavingTime () AS BOOLEAN
   DIM tzi AS TIME_ZONE_INFORMATION
   DIM r AS DWORD = GetTimeZoneInformation(@tzi)
   RETURN (r = TIME_ZONE_ID_DAYLIGHT)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Current bias for local time translation. The bias is the difference between Coordinated
' Universal Time (UTC) and local time. All translations between UTC and local time are
' based on the following formula: UTC = local time + bias. Units = minutes.
' ========================================================================================
PRIVATE FUNCTION AfxTimeZoneBias () AS DWORD
   DIM tzi AS TIME_ZONE_INFORMATION
   DIM r AS DWORD = GetTimeZoneInformation(@tzi)
   IF r = TIME_ZONE_ID_INVALID THEN RETURN 0
   RETURN tzi.Bias
END FUNCTION
' ========================================================================================

' ========================================================================================
' Bias value to be used during local time translations that occur during daylight saving
' time. This property is ignored if a value for the DaylightDay property is not supplied.
' The value of this property is added to the Bias property to form the bias used during
' daylight time. In most time zones, the value of this property is 60. Units = minutes.
' ========================================================================================
PRIVATE FUNCTION AfxTimeZoneDaylightBias () AS DWORD
   DIM tzi AS TIME_ZONE_INFORMATION
   DIM r AS DWORD = GetTimeZoneInformation(@tzi)
   IF r = TIME_ZONE_ID_INVALID THEN RETURN 0
   RETURN tzi.DaylightBias
END FUNCTION
' ========================================================================================

' ========================================================================================
' A description for standard time. For example, "EST" could indicate Eastern Standard Time.
' The string will be returned unchanged by the GetTimeZoneInformation function.
' This string can be empty.
' ========================================================================================
PRIVATE FUNCTION AfxTimeZoneStandardName () AS CWSTR
   DIM tzi AS TIME_ZONE_INFORMATION
   DIM r AS DWORD = GetTimeZoneInformation(@tzi)
   IF r = TIME_ZONE_ID_INVALID THEN RETURN ""
   RETURN tzi.StandardName
END FUNCTION
' ========================================================================================

' ========================================================================================
' A description for daylight saving time. For example, "PDT" could indicate Pacific
' Daylight Time. The string will be returned unchanged by the GetTimeZoneInformation
' function. This string can be empty.
' ========================================================================================
PRIVATE FUNCTION AfxTimeZoneDaylightName () AS CWSTR
   DIM tzi AS TIME_ZONE_INFORMATION
   DIM r AS DWORD = GetTimeZoneInformation(@tzi)
   IF r = TIME_ZONE_ID_INVALID THEN RETURN ""
   RETURN tzi.DaylightName
END FUNCTION
' ========================================================================================

' ========================================================================================
' DaylightDayOfWeek of the DaylightMonth when the transition from standard time to daylight
' saving time occurs on this operating system.
' Example: If the transition day (DaylightDayOfWeek) occurs on a Sunday, then the value
' "1" indicates the first Sunday of the DaylightMonth, "2" indicates the second Sunday,
' and so on. The value "5" indicates the last DaylightDayOfWeek in the month.
' ========================================================================================
PRIVATE FUNCTION AfxTimeZoneDaylightDay () AS WORD
   DIM tzi AS TIME_ZONE_INFORMATION
   DIM r AS DWORD = GetTimeZoneInformation(@tzi)
   IF r = TIME_ZONE_ID_INVALID THEN RETURN 0
   RETURN tzi.DaylightDate.wDay
END FUNCTION
' ========================================================================================

' ========================================================================================
' Day of the week when the transition from standard time to daylight saving time occurs on
' an operating system. 0 = Sunday, 1 = Monday, and so on.
' ========================================================================================
PRIVATE FUNCTION AfxTimeZoneDaylightDayOfWeek () AS WORD
   DIM tzi AS TIME_ZONE_INFORMATION
   DIM r AS DWORD = GetTimeZoneInformation(@tzi)
   IF r = TIME_ZONE_ID_INVALID THEN RETURN 0
   RETURN tzi.DaylightDate.wDayOfWeek
END FUNCTION
' ========================================================================================

' ========================================================================================
' Hour of the day when the transition from standard time to daylight saving time occurs on
' an operating system.
' ========================================================================================
PRIVATE FUNCTION AfxTimeZoneDaylightHour () AS WORD
   DIM tzi AS TIME_ZONE_INFORMATION
   DIM r AS DWORD = GetTimeZoneInformation(@tzi)
   IF r = TIME_ZONE_ID_INVALID THEN RETURN 0
   RETURN tzi.DaylightDate.wHour
END FUNCTION
' ========================================================================================

' ========================================================================================
' Month when the transition from standard time to daylight saving time occurs on an
' operating system.
' ========================================================================================
PRIVATE FUNCTION AfxTimeZoneDaylightMonth () AS WORD
   DIM tzi AS TIME_ZONE_INFORMATION
   DIM r AS DWORD = GetTimeZoneInformation(@tzi)
   IF r = TIME_ZONE_ID_INVALID THEN RETURN 0
   RETURN tzi.DaylightDate.wMonth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the local date and time as a DATE_ (double).
' ========================================================================================
PRIVATE FUNCTION AfxLocalVariantTime () AS DATE_
   DIM ST AS SYSTEMTIME, dt AS DATE_
   GetLocalTime @ST
   SystemTimeToVariantTime @ST, @dt
   RETURN dt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a SYSTEMTIME to a DATE_ (double).
' ========================================================================================
PRIVATE FUNCTION AfxSystemTimeToVariantTime (BYREF ST AS SYSTEMTIME) AS DATE_
   DIM dt AS DATE_
   SystemTimeToVariantTime @ST, @dt
   RETURN dt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a DATE_ (double) to a SYSTEMTIME.
' ========================================================================================
PRIVATE FUNCTION AfxVariantTimeToSystemTime (BYVAL dt AS DATE_) AS SYSTEMTIME
   DIM ST AS SYSTEMTIME
   VariantTimeToSystemTime dt, @ST
   RETURN ST
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a FILETIME to a DATE_ (double).
' ========================================================================================
PRIVATE FUNCTION AfxFileTimeToVariantTime (BYREF FT AS FILETIME) AS DATE_
   DIM dt AS DATE_, ST AS SYSTEMTIME
   FileTimeToSystemTime(@FT, @ST)
   SystemTimeToVariantTime @ST, @dt
   RETURN dt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a DATE_ type to a string containing the date and the time.
' ========================================================================================
PRIVATE FUNCTION AfxVariantDateTimeToStr (BYVAL vbDate AS DATE_, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT, BYVAL dwFlags AS DWORD = 0) AS CWSTR
   DIM bstrOut AS AFX_BSTR
   VarBstrFromDate(vbDate, lcid, dwFlags, @bstrOut)
   DIM cwsOut AS CWSTR = *bstrOut
   SysFreeString bstrOut
   RETURN cwsOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a DATE_ type to a string containing only the date.
' ========================================================================================
PRIVATE FUNCTION AfxVariantDateToStr (BYVAL vbDate AS DATE_, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM bstrOut AS AFX_BSTR
   VarBstrFromDate(vbDate, lcid, VAR_DATEVALUEONLY, @bstrOut)
   DIM cwsOut AS CWSTR = *bstrOut
   SysFreeString bstrOut
   RETURN cwsOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a DATE_ type to a string containing only the time.
' ========================================================================================
PRIVATE FUNCTION AfxVariantTimeToStr (BYVAL vbDate AS DATE_, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM bstrOut AS AFX_BSTR
   VarBstrFromDate(vbDate, lcid, VAR_TIMEVALUEONLY, @bstrOut)
   DIM cwsOut AS CWSTR = *bstrOut
   SysFreeString bstrOut
   RETURN cwsOut
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a FILETIME type to a string containing the date based on the specified mask, e.g. "dd-MM-yyyy"
' ========================================================================================
PRIVATE FUNCTION AfxFileTimeToDateStr (BYREF FT AS FILETIME, BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, wszDateStr AS WSTRING * 260
   FileTimeToSystemTime(@FT, @ST)
   GetDateFormatW(lcid, NULL, @ST, wszMask, wszDateStr, SIZEOF(wszDateStr))
   RETURN wszDateStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a FILETIME type to a string containing the time based on the specified mask, e.g. "hh':'mm':'ss".
' ========================================================================================
PRIVATE FUNCTION AfxFileTimeToTimeStr (BYREF FT AS FILETIME, BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM ST AS SYSTEMTIME, wszDateStr AS WSTRING * 260
   FileTimeToSystemTime(@FT, @ST)
   GetTimeFormatW(lcid, NULL, @ST, wszMask, wszDateStr, SIZEOF(wszDateStr))
   RETURN wszDateStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a SYSTEMTIME type to a string containing the date and the time. based on the
' specified mask, e.g. "dd-MM-yyyy"
' ========================================================================================
PRIVATE FUNCTION AfxSystemTimeToDateStr (BYREF ST AS SYSTEMTIME, BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszDateStr AS WSTRING * 260
   GetDateFormatW(lcid, NULL, @ST, wszMask, wszDateStr, SIZEOF(wszDateStr))
   RETURN wszDateStr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a FILETIME type to a string containing the date and the time. based on the
' specified mask, e.g. "hh':'mm':'ss".
' ========================================================================================
PRIVATE FUNCTION AfxSystemTimeToTimeStr (BYREF ST AS SYSTEMTIME, BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
   DIM wszDateStr AS WSTRING * 260
   GetTimeFormatW(lcid, NULL, @ST, wszMask, wszDateStr, SIZEOF(wszDateStr))
   RETURN wszDateStr
END FUNCTION
' ========================================================================================
