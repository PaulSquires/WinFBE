' ########################################################################################
' Microsoft Windows
' File: CTime64.inc
' Contents: Date and time classes
' Compiler: FreeBasic 32 & 64-bit
' Based on the MFC classes CTimeSpan and CTime.
' Copyright (C) Microsoft Corporation
' Written in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "Afx/AfxTime.inc"
USING Afx

NAMESPACE Afx

' ########################################################################################
'                                *** CTimeSpan Class ***
' ########################################################################################

' ========================================================================================
' An amount of time, which is internally stored as the number of seconds in the time span.
' ========================================================================================
TYPE CTimeSpan

Public:
   m_timeSpan AS LONGLONG

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL lltime AS LONGLONG)
   DECLARE CONSTRUCTOR (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS LONGLONG
   DECLARE OPERATOR LET (BYVAL nSpan AS LONGLONG)
   DECLARE OPERATOR LET (BYREF cSpan AS CTimeSpan)
   DECLARE OPERATOR += (BYREF cSpan AS CTimeSpan)
   DECLARE OPERATOR -= (BYREF cSpan AS CTimeSpan)
   DECLARE SUB SetTimeSpan (BYVAL lltime AS LONGLONG)
   DECLARE SUB SetTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   DECLARE FUNCTION GetTimeSpan () AS LONGLONG
   DECLARE FUNCTION GetTotalHours () AS LONGLONG
   DECLARE FUNCTION GetTotalMinutes () AS LONGLONG
   DECLARE FUNCTION GetTotalSeconds () AS LONGLONG
   DECLARE FUNCTION GetDays () AS LONG
   DECLARE FUNCTION GetHours () AS LONG
   DECLARE FUNCTION GetMinutes () AS LONG
   DECLARE FUNCTION GetSeconds () AS LONG

END TYPE
' ========================================================================================

' ========================================================================================
' CTimeSpan constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CTimeSpan
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' - lltime: A time value, which is the number of seconds in the time span.
' ========================================================================================
PRIVATE CONSTRUCTOR CTimeSpan (BYVAL lltime AS LONGLONG)
   m_timeSpan = lltime
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' - lDays, nHours, nMins, nSecs
' Indicate the day and time values to be copied into the new CTimeSpan object.
' ========================================================================================
PRIVATE CONSTRUCTOR CTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   m_timeSpan = nSecs + 60* (nMins + 60* (nHours + 24 * lDays))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' CTimeSpan destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTimeSpan
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the CTimeSpan value.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.CAST () AS LONGLONG
   RETURN m_timeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Sets the value of this date/time-span value.
' - lltime: A time value, which is the number of seconds in the time span.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.LET (BYVAL lltime AS LONGLONG)
   m_timeSpan = lltime
END OPERATOR
' ========================================================================================
' ========================================================================================
' Sets the value of this date/time-span value.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.LET (BYREF cSpan AS CTimeSpan)
   m_timeSpan = cSpan.m_timeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Sets the value of this date/time-span value.
' - lltime: A time value, which is the number of seconds in the time span.
' ========================================================================================
PRIVATE SUB CTimeSpan.SetTimeSpan (BYVAL lltime AS LONGLONG)
   m_timeSpan = lltime
END SUB
' ========================================================================================
' ========================================================================================
' Sets the value of this date/time-span value.
' ========================================================================================
PRIVATE SUB CTimeSpan.SetTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
   m_timeSpan = nSecs + 60* (nMins + 60* (nHours + 24 * lDays))
END SUB
' ========================================================================================
' ========================================================================================
' Returns the value of the CTimeSpan object.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetTimeSpan () AS LONGLONG
   RETURN m_timeSpan
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a CTimeSpan object to this CTimeSpan.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.+= (BYREF cSpan AS CTimeSpan)
   m_timeSpan += cSpan.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
' Subtracts a CTimeSpan object to from this CTimeSpan.
' ========================================================================================
PRIVATE OPERATOR CTimeSpan.-= (BYREF cSpan AS CTimeSpan)
   m_timeSpan -= cSpan.m_timeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a value that represents the total number of complete hours in this CTimeSpan.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetTotalHours () AS LONGLONG
   RETURN m_timeSpan/3600
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a value that represents the total number of complete minutes in this CTimeSpan.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetTotalMinutes () AS LONGLONG
   RETURN m_timeSpan/60
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves this date/time-span value expressed in seconds.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetTotalSeconds () AS LONGLONG
   RETURN m_timeSpan
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a value that represents the number of days in this CTimeSpan.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetDays () AS LONG
   RETURN m_timeSpan/(24*3600) 
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a value that represents the number of hours in the current day (–23 through 23).
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetHours () AS LONG
   RETURN this.GetTotalHours()-(this.GetDays()*24)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns a value that represents the number of minutes in the current hour (–59 through 59).
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetMinutes () AS LONG
   RETURN this.GetTotalMinutes()-(this.GetTotalHours()*60)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Retrieves the seconds portion of this date/time-span value.
' The return values from this function range between – 59 and 59.
' ========================================================================================
PRIVATE FUNCTION CTimeSpan.GetSeconds () AS LONG
   RETURN this.GetTotalSeconds()-(this.GetTotalMinutes()*60)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                 *** CTime64 Class ***
' ########################################################################################

' ========================================================================================
' Represents an absolute time and date.
' Note: Named CTime64 because CTime conflicts with the C function ctime.
' ========================================================================================
TYPE CTime64

Public:
   m_time AS LONGLONG

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL timeSrc AS LONGLONG)
   DECLARE CONSTRUCTOR (BYVAL nYear AS LONG, BYVAL nMonth AS LONG, BYVAL nDay AS LONG, BYVAL nHour AS LONG, BYVAL nMin AS LONG, BYVAL nSec AS LONG)
   DECLARE CONSTRUCTOR (BYREF systimeSrc AS SYSTEMTIME)
   DECLARE CONSTRUCTOR (BYREF filetimeSrc AS FILETIME)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS LONGLONG
   DECLARE OPERATOR LET (BYVAL timeSrc AS LONGLONG)
   DECLARE OPERATOR LET (BYREF st AS SYSTEMTIME)
   DECLARE OPERATOR LET (BYREF ft AS FILETIME)
   DECLARE OPERATOR += (BYREF dateSpan AS CTimeSpan)
   DECLARE OPERATOR -= (BYREF dateSpan AS CTimeSpan)
   DECLARE FUNCTION SetDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, BYVAL nHour AS WORD = 0, BYVAL nMin AS WORD = 0, BYVAL nSec AS WORD = 0) AS BOOLEAN
   DECLARE FUNCTION GetCurrentTime () AS CTime64
   DECLARE FUNCTION GetTime() AS LONGLONG
   DECLARE FUNCTION GetYear () AS LONG
   DECLARE FUNCTION GetMonth () AS LONG
   DECLARE FUNCTION GetDay () AS LONG
   DECLARE FUNCTION GetHour () AS LONG
   DECLARE FUNCTION GetMinute () AS LONG
   DECLARE FUNCTION GetSecond () AS LONG
   DECLARE FUNCTION GetDayOfWeek () AS LONG
   DECLARE FUNCTION GetGmtTime () AS tm
   DECLARE FUNCTION GetLocalTime () AS tm
   DECLARE FUNCTION GetAsSystemTime () AS SYSTEMTIME
   DECLARE FUNCTION GetAsFileTime () AS FILETIME
   DECLARE FUNCTION Format (BYREF wszFmt AS WSTRING) AS CWSTR
   DECLARE FUNCTION FormatGmt (BYREF wszFmt AS WSTRING) AS CWSTR

END TYPE
' ========================================================================================

' ========================================================================================
' CTime64 constructors
' All these constructors create new CTime64 objects initialized to the specified value.
' - timeSrc : Indicates __time64_t (LONGLONG) value.
' - systimeSrc: A SYSTEMTIME structure to be converted to a __time64_t (LONGLONG) value
'   and copied into the new CTime64 object.
' - filetimeSrc: A FILETIME structure to be converted to a __time64_t (LONGLONG) value
'   and copied into  the new CTime64 object. Note that FILETIME uses Universal Coordinated
'   Time (UTC), so if you pass a local time in the structure, your results will be incorrect.
' - nYear, nMonth, nDay, nHour, nMin, nSec
'   Indicates the date and time values to be copied into the new CTime64 object.
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' CTime64 destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTime64
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' * Initializes a CTime64 object from an existing __time64_t value.
' Example:
' DIM ct1 AS CTime64 = CTime64().GetCurrentTime()
' DIM ct2 AS CTime64 = ct1.GetTime
' Note: You can also use DIM ct2 AS CTime64 = ct1
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64 (BYVAL timeSrc AS LONGLONG)
   m_time = timeSrc
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' * Initializes a CTime64 object from individual date/time values.
' Parameters:
' - nYear  : The year
' - nMonth : The month
' - nDay   : The day
' - nHour  : The hour
' - nMin   : The minutes
' - nSec   : The seconds
' Usage example:
' ' // Year = 2017, Month = 10 (October), Day = 9, Hour = 11, Minutes = 32, Seconds = 45
' DIM ct AS CTime64 = CTime64(2017, 10, 9, 11, 32, 45)
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64 (BYVAL nYear AS LONG, BYVAL nMonth AS LONG, BYVAL nDay AS LONG, BYVAL nHour AS LONG, BYVAL nMin AS LONG, BYVAL nSec AS LONG)
	IF nYear < 1900 THEN EXIT CONSTRUCTOR
	IF nMonth < 1 OR nMonth > 12 THEN EXIT CONSTRUCTOR
	IF nDay < 1 OR nDay >= 31 THEN EXIT CONSTRUCTOR
	IF nHour < 0 OR nHour > 23 THEN EXIT CONSTRUCTOR
	IF nMin < 0 OR nMin > 59 THEN EXIT CONSTRUCTOR
	IF nSec < 0 OR nSec > 59 THEN EXIT CONSTRUCTOR
   DIM atm AS tm
   atm.tm_year = nYear - 1900   ' // tm_year is 1900 based
   atm.tm_mon = nMonth - 1      ' // tm_mon is 0 based
   atm.tm_mday = nDay
   atm.tm_hour = nHour
   atm.tm_min = nMin
   atm.tm_sec = nSec
   m_time = AfxMakeTime64(atm)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' * Initializes a CTime64 object from a SYSTEMTIME structure.
' Example:
' ' // Year = 2017, Month = 10 (October), Day = 9, Hour = 11, Minutes = 32, Seconds = 45
' DIM st AS SYSTEMTIME 
' st.wYear = 2017
' st.wMonth = 10
' st.wDayOfWeek = 0
' st.wDay = 9
' st.wHour = 11
' st.wMinute = 32
' st.wSecond = 45
' st.wMilliseconds = 0
' --or--
' DIM st AS SYSTEMTIME = (2017, 10, 0, 9, 11, 32, 45, 0)
' DIM ct AS CTime64 = st
' --or--
' DIM ct AS CTime64
' ct.SetDateTime(2017, 10, 9, 11, 32, 45)
' DIM st AS SYSTEMTIME = ct.GetAsSystemtime
' DIM ct2 AS CTime64 = st
' --or--
' DIM ct AS CTime64
' ct.SetDateTime(2017, 10, 9, 11, 32, 45)
' DIM ct2 AS CTime64 = ct.GetAsSystemtime
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64 (BYREF systimeSrc AS SYSTEMTIME)
   this = CTime64(systimeSrc.wYear, systimeSrc.wMonth, systimeSrc.wDay, _
          systimeSrc.wHour, systimeSrc.wMinute, systimeSrc.wSecond)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Constructs a CTime64 object from a FILETIME structure.
' This constructor assumes a time based on UTC and automatically converts the value to
' local time before storing the result.
' ========================================================================================
PRIVATE CONSTRUCTOR CTime64 (BYREF filetimeSrc AS FILETIME)
   ' // first convert file time (UTC time) to local time
   DIM ftlocalTime AS FILETIME
   IF FileTimeToLocalFileTime(@filetimeSrc, @ftlocalTime) = 0 THEN EXIT CONSTRUCTOR
   ' // then convert that time to system time
   DIM sysTime AS SYSTEMTIME
   IF FileTimeToSystemTime(@ftlocalTime, @sysTime) = 0 THEN EXIT CONSTRUCTOR
   this = CTime64(systime.wYear, systime.wMonth, systime.wDay, systime.wHour, _
          systime.wMinute, systime.wSecond)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Returns the underlying m_time value from this CTime64 object.
' ========================================================================================
PRIVATE OPERATOR CTime64.CAST () AS LONGLONG
   RETURN m_time
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a __time64_t (LONGLONG) value to this CTime64 object.
' ========================================================================================
PRIVATE OPERATOR CTime64.LET (BYVAL timeSrc AS LONGLONG)
   m_time = timeSrc
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a SYSTEMTIME structure.
' Example:
' DIM ct AS CTime64
' ct.SetDateTime(2017, 10, 9, 11, 32, 45)
' DIM st AS SYSTEMTIME = ct.GetAsSystemtime
' DIM ct2 AS CTime64
' ct2 = st
' ========================================================================================
PRIVATE OPERATOR CTime64.LET (BYREF st AS SYSTEMTIME)
   this = CTime64(st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a FILETIME structure.
' ========================================================================================
PRIVATE OPERATOR CTime64.LET (BYREF ft AS FILETIME)
   ' // first convert file time (UTC time) to local time
   DIM ftlocalTime AS FILETIME
   IF FileTimeToLocalFileTime(@ft, @ftlocalTime) = 0 THEN EXIT OPERATOR
   ' // then convert that time to system time
   DIM sysTime AS SYSTEMTIME
   IF FileTimeToSystemTime(@ftlocalTime, @sysTime) = 0 THEN EXIT OPERATOR
   this = CTime64(systime.wYear, systime.wMonth, systime.wDay, systime.wHour, _
          systime.wMinute, systime.wSecond)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Adds a CTimeSpan value to this CTime64 object.
' ========================================================================================
PRIVATE OPERATOR CTime64.+= (BYREF dateSpan AS CTimeSpan)
   m_time += dateSpan.GetTimeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CTimeSpan value from this CTime64 object.
' ========================================================================================
PRIVATE OPERATOR CTime64.-= (BYREF dateSpan AS CTimeSpan)
   m_time -= dateSpan.GetTimeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Sets the date and time of this CTime64 object.
' Parameters:
' - nYear  : The year (1970-3000)
' - nMonth : The month (1-12)
' - nDay   : The day (1-31)
' - nHour  : The hour (0-23)
' - nMin   : The minutes (0-59)
' - nSec   : The seconds (0-59)
' Example:
' ' // Year = 2017, Month = 10 (October), Day = 9, Hour = 11, Minutes = 32, Seconds = 45
' DIM ct AS CTime64
' ct.SetDateTime(2017, 10, 9, 11, 32, 45)
' ========================================================================================
PRIVATE FUNCTION CTime64.SetDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, BYVAL nHour AS WORD = 0, BYVAL nMin AS WORD = 0, BYVAL nSec AS WORD = 0) AS BOOLEAN
	IF nYear < 1900 THEN RETURN FALSE
	IF nMonth < 1 OR nMonth > 12 THEN RETURN FALSE
	IF nDay < 1 OR nDay >= 31 THEN RETURN FALSE
	IF nHour > 23 THEN RETURN FALSE
	IF nMin > 59 THEN RETURN FALSE
	IF nSec > 59 THEN RETURN FALSE
   DIM atm AS tm
   atm.tm_year = nYear - 1900   ' // tm_year is 1900 based
   atm.tm_mon = nMonth - 1      ' // tm_mon is 0 based
   atm.tm_mday = nDay
   atm.tm_hour = nHour
   atm.tm_min = nMin
   atm.tm_sec = nSec
   m_time = AfxMakeTime64(atm)
   RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a CTime64 object that represents the current time.
' Usage example: DIM ct AS CTime64 = CTime64().GetCurrentTime()
' --or--
' DIM ct AS CTime64
' ct = ct.GetCurrentTime()
' Usage example:
' DIM ct1 AS CTime64 = CTime64().GetCurrentTime()
' DIM ct2 AS CTime64 = ct1 + CTimeSpan(0, 1, 0, 0)   ' // 1 hour later
' Note: The use of () after GetCurrentTime is needed with FreeBasic; otherwise, the
' compiler returns an error 18: Element not defined.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetCurrentTime () AS CTime64
   RETURN CTime64(AfxTime64)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a __time64_t (LONGLONG) value for the given CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetTime() AS LONGLONG
   RETURN m_time
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the year represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetYear() AS LONG
   DIM _tm AS tm = AfxLocalTime64(m_time)
   RETURN _tm.tm_year + 1900
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the month represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetMonth() AS LONG
   DIM _tm AS tm = AfxLocalTime64(m_time)
   RETURN _tm.tm_mon + 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the day represent by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetDay() AS LONG
   DIM _tm AS tm = AfxLocalTime64(m_time)
   RETURN _tm.tm_mday
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the hour represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetHour() AS LONG
   DIM _tm AS tm = AfxLocalTime64(m_time)
   RETURN _tm.tm_hour
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the minute represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetMinute() AS LONG
   DIM _tm AS tm = AfxLocalTime64(m_time)
   RETURN _tm.tm_min
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the second represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetSecond() AS LONG
   DIM _tm AS tm = AfxLocalTime64(m_time)
   RETURN _tm.tm_sec
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the day of the week represented by the CTime64 object.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetDayOfWeek () AS LONG
   DIM _tm AS tm = AfxLocalTime64(m_time)
   RETURN _tm.tm_wday + 1
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a tm structure that contains a decomposition of the time contained in this CTime64 object.
' The fields of the structure type tm store the following values, each of which is an int:
' tm_sec : Seconds after minute (0 – 59).
' tm_min : Minutes after hour (0 – 59).
' tm_hour : Hours after midnight (0 – 23).
' tm_mday : Day of month (1 – 31).
' tm_mon : Month (0 – 11; January = 0).
' tm_year : Year (current year minus 1900).
' tm_wday : Day of week (0 – 6; Sunday = 0).
' tm_yday : Day of year (0 – 365; January 1 = 0).
' tm_isdst : Positive value if daylight saving time is in effect; 0 if daylight saving
' time is not in effect; negative value if status of daylight saving time is unknown.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetGmtTime () AS tm
   RETURN AfxGmtTime64(m_time)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a tm  structure containing a decomposition of the time contained in this CTime64 object.
' The fields of the structure type tm store the following values, each of which is an int:
' tm_sec : Seconds after minute (0 – 59).
' tm_min : Minutes after hour (0 – 59).
' tm_hour : Hours after midnight (0 – 23).
' tm_mday : Day of month (1 – 31).
' tm_mon : Month (0 – 11; January = 0).
' tm_year : Year (current year minus 1900).
' tm_wday : Day of week (0 – 6; Sunday = 0).
' tm_yday : Day of year (0 – 365; January 1 = 0).
' tm_isdst : Positive value if daylight saving time is in effect; 0 if daylight saving
' time is not in effect; negative value if status of daylight saving time is unknown.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetLocalTime () AS tm
   RETURN AfxLocalTime64(m_time)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the time information stored in the CTime64 object to a Win32–compatible
' SYSTEMTIME structure.
' ========================================================================================
PRIVATE FUNCTION CTime64.GetAsSystemTime () AS SYSTEMTIME
   DIM ttm AS tm = this.GetLocalTime
   DIM ts AS SYSTEMTIME
   ts.wYear = 1900 + ttm.tm_year
	ts.wMonth = 1 + ttm.tm_mon
	ts.wDayOfWeek = ttm.tm_wday
	ts.wDay = ttm.tm_mday
	ts.wHour = ttm.tm_hour
	ts.wMinute = ttm.tm_min
	ts.wSecond = ttm.tm_sec
	ts.wMilliseconds = 0
   RETURN ts
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts the time information stored in the CTime64 object to a Win32–compatible
' FILETIME structure. A FILETIME structure contains a 64-bit value representing the number
' of 100-nanosecond intervals since January 1, 1601 (UTC).
' ========================================================================================
PRIVATE FUNCTION CTime64.GetAsFileTime () AS FILETIME
   DIM ft AS FILETIME
   DIM st AS SYSTEMTIME = this.GetAsSystemTime
   SystemTimeToFileTime(@st, @ft)
   RETURN ft
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a CTime64 object to a string.
' This form formats the value by using the format string which contains special formatting
' codes that are preceded by a percent sign (%), as in printf.
' For more information about the formatting codes, see strftime, wcsftime in the Run-Time
' Library Reference: https://msdn.microsoft.com/en-us/library/38wh24td.aspx#coledatetime__format
' Example: print ct.Format("%A, %B %d, %Y %H:%M:%S")
' Example: print ct.FormatGmt("%A, %B %d, %Y %H:%M:%S")
' ========================================================================================
PRIVATE FUNCTION CTime64.Format (BYREF wszFmt AS WSTRING) AS CWSTR
   DIM cwsOut AS CWSTR
   IF LEN(wszFmt) = 0 THEN RETURN cwsOut
   cwsOut = STRING(128, CHR(0))
   DIM tmTemp AS tm = AfxLocalTime64(m_time)
   wcsftime(cwsOut.vptr, 128, @wszFmt, @tmTemp)
   RETURN cwsOut
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTime64.FormatGmt (BYREF wszFmt AS WSTRING) AS CWSTR
   DIM cwsOut AS CWSTR
   IF LEN(wszFmt) = 0 THEN RETURN cwsOut
   cwsOut = STRING(128, CHR(0))
   DIM tmTemp AS tm = AfxGmtTime64(m_time)
   wcsftime(cwsOut.vptr, 128, @wszFmt, @tmTemp)
   RETURN cwsOut
END FUNCTION
' ========================================================================================

END NAMESPACE

' ########################################################################################
'                                *** Global operators ***
' ########################################################################################

USING Afx

' ========================================================================================
' Adds CTimeSpan values.
' ========================================================================================
PRIVATE OPERATOR + (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS CTimeSpan
   OPERATOR = CTimeSpan(cSpan1.m_timeSpan + cSpan2.m_timeSpan )
END OPERATOR
' ========================================================================================
' ========================================================================================
' Subtracts CTimeSpan values.
' Note: Can also be used to find the difference between two dates. Subtract one date
' from another and call the CTimeSpan GetDays, GetHours, GetMinutes or GetSeconds methods.
' ========================================================================================
PRIVATE OPERATOR - (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS CTimeSpan
   OPERATOR = CTimeSpan(cSpan1.m_timeSpan - cSpan2.m_timeSpan )
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators
' ========================================================================================
PRIVATE OPERATOR = (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan = cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan <> cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan < cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan > cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan <= cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_timeSpan >= cSpan2.m_timeSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Adds a CTimeSpan value to a CTime64 value
' ========================================================================================
PRIVATE OPERATOR + (BYREF dt AS CTime64, BYREF dateSpan AS CTimeSpan) AS CTime64
   RETURN CTime64(dt.m_time + dateSpan.GetTimeSpan)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CTimeSpan value from a CTime64 value
' ========================================================================================
PRIVATE OPERATOR - (BYREF dt AS CTime64, BYREF dateSpan AS CTimeSpan) AS CTime64
   RETURN CTime64(dt.m_time - dateSpan.GetTimeSpan)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CTime64 value from another CTime64 value
' ========================================================================================
PRIVATE OPERATOR - (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS CTimeSpan
   RETURN CTimeSpan(dt1.m_time - dt2.m_time)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators
' ========================================================================================
PRIVATE OPERATOR = (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time = dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time <> dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time < dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time > dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time <= dt2.m_time
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
   RETURN dt1.m_time >= dt2.m_time
END OPERATOR
' ========================================================================================
