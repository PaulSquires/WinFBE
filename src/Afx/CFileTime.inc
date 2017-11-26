' ########################################################################################
' Microsoft Windows
' File: CFileTime.inc
' Contents: Date and time classes
' Compiler: FreeBasic 32 & 64-bit
' Based on the MFC classes CFileTimeSpan and CFileTime.
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

CONST CFileTime_Millisecond = 10000ull
CONST CFileTime_Second = CFileTime_Millisecond * 1000
CONST CFileTime_Minute = CFileTime_Second * 60
CONST CFileTime_Hour = CFileTime_Minute * 60
CONST CFileTime_Day = CFileTime_Hour * 24
CONST CFileTime_Week = CFileTime_Day * 7

' ########################################################################################
'                                *** CFileTimeSpan Class ***
' ########################################################################################

' ========================================================================================
' This class provides methods for managing relative date and time values associated with a file.
' ========================================================================================
TYPE CFileTimeSpan

Public:
   m_nSpan AS LONGLONG

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL nSpan AS LONGLONG)
   DECLARE CONSTRUCTOR (BYREF cSpan AS CFileTimeSpan)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS LONGLONG
   DECLARE OPERATOR LET (BYVAL nSpan AS LONGLONG)
   DECLARE OPERATOR LET (BYREF cSpan AS CFileTimeSpan)
   DECLARE SUB SetTimeSpan (BYVAL nSpan AS LONGLONG)
   DECLARE FUNCTION GetTimeSpan () AS LONGLONG
   DECLARE OPERATOR += (BYREF cSpan AS CFileTimeSpan)
   DECLARE OPERATOR -= (BYREF cSpan AS CFileTimeSpan)
   DECLARE FUNCTION Day () AS ULONGLONG
   DECLARE FUNCTION Hour () AS ULONGLONG
   DECLARE FUNCTION Minute () AS ULONGLONG
   DECLARE FUNCTION Second () AS ULONGLONG
   DECLARE FUNCTION MilliSecond () AS ULONGLONG
   DECLARE FUNCTION Week () AS ULONGLONG

END TYPE
' ========================================================================================

' ========================================================================================
' CFileTimeSpan constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CFileTimeSpan
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' - nSpan: A period of time in milliseconds.
' ========================================================================================
PRIVATE CONSTRUCTOR CFileTimeSpan (BYVAL nSpan AS LONGLONG)
   m_nSpan = nSpan
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' - cSpan: A CFileTimeSpan object.
' ========================================================================================
PRIVATE CONSTRUCTOR CFileTimeSpan (BYREF cSpan AS CFileTimeSpan)
   m_nSpan = cSpan.m_nSpan
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' CFileTimeSpan destructor
' ========================================================================================
PRIVATE DESTRUCTOR CFileTimeSpan
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the CFileTimeSpan value.
' ========================================================================================
PRIVATE OPERATOR CFileTimeSpan.CAST () AS LONGLONG
   RETURN m_nSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Sets the time span of the CFileTimeSpan object.
' - nSpan: The new value for the time span in milliseconds.
' ========================================================================================
PRIVATE OPERATOR CFileTimeSpan.LET (BYVAL nSpan AS LONGLONG)
   m_nSpan = nspan
END OPERATOR
' ========================================================================================
' ========================================================================================
' Sets the time span of the CFileTimeSpan object.
' - cSpan : A CFileTimeSpan object.
' ========================================================================================
PRIVATE OPERATOR CFileTimeSpan.LET (BYREF cSpan AS CFileTimeSpan)
   m_nSpan = cSpan.m_nSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Sets the time span of the CFileTimeSpan object.
' - nSpan: The new value for the time span in milliseconds.
' ========================================================================================
PRIVATE SUB CFileTimeSpan.SetTimeSpan (BYVAL nspan AS LONGLONG)
   m_nSpan = nSpan
END SUB
' ========================================================================================

' ========================================================================================
' Returns the value of the CFileTimeSpan object.
' ========================================================================================
PRIVATE FUNCTION CFileTimeSpan.GetTimeSpan () AS LONGLONG
   RETURN m_nSpan
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a CFileTimeSpan object to this CFileTimeSpan.
' ========================================================================================
PRIVATE OPERATOR CFileTimeSpan.+= (BYREF cSpan AS CFileTimeSpan)
   m_nSpan += cSpan.m_nSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CFileTimeSpan object to from this CFileTimeSpan.
' ========================================================================================
PRIVATE OPERATOR CFileTimeSpan.-= (BYREF cSpan AS CFileTimeSpan)
   m_nSpan -= cSpan.m_nSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' The number of 100-nanosecond intervals that make up one day.
' ========================================================================================
PRIVATE FUNCTION CFileTimeSpan.Day () AS ULONGLONG
   RETURN CFileTime_Day
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one hour.
' ========================================================================================
PRIVATE FUNCTION CFileTimeSpan.Hour () AS ULONGLONG
   RETURN CFileTime_Hour
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one minute.
' ========================================================================================
PRIVATE FUNCTION CFileTimeSpan.Minute () AS ULONGLONG
   RETURN CFileTime_Minute
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one second.
' ========================================================================================
PRIVATE FUNCTION CFileTimeSpan.Second () AS ULONGLONG
   RETURN CFileTime_Second
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one millisecond.
' ========================================================================================
PRIVATE FUNCTION CFileTimeSpan.MilliSecond () AS ULONGLONG
   RETURN CFileTime_Millisecond
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one week.
' ========================================================================================
PRIVATE FUNCTION CFileTimeSpan.Week () AS ULONGLONG
   RETURN CFileTime_Week
END FUNCTION
' ========================================================================================

' ########################################################################################
'                                 *** CFileTime Class ***
' ########################################################################################

' ========================================================================================
' This class provides methods for managing the date and time values associated with the
' creation, access and modification of files. The methods and data of this class are
' frequently used in conjunction with CFileTimeSpan objects, which deal with relative time
' values. The date and time value is stored as a 64-bit value representing the number of
' 100-nanosecond intervals since January 1, 1601. This is the Coordinated Universal Time
' (UTC) format.
' Note Not all file systems can record creation and last access time and not all file
' systems record them in the same manner. For example, on the Windows NT FAT file system,
' create time has a resolution of 10 milliseconds, write time has a resolution of 2 seconds,
' and access time has a resolution of 1 day (the access date). On NTFS, access time has a
' resolution of 1 hour. Furthermore, FAT records times on disk in local time, but NTFS
' records times on disk in UTC.
' ========================================================================================
TYPE CFileTime

Public:
	dwLowDateTime AS DWORD
	dwHighDateTime AS DWORD

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL nTime AS ULONGLONG)
   DECLARE CONSTRUCTOR (BYREF ft AS FILETIME)
   DECLARE CONSTRUCTOR (BYREF st AS SYSTEMTIME)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS LONGLONG
   DECLARE OPERATOR LET (BYVAL nTime AS ULONGLONG)
   DECLARE OPERATOR LET (BYREF ft AS FILETIME)
   DECLARE OPERATOR LET (BYREF st AS SYSTEMTIME)
   DECLARE OPERATOR += (BYREF dateSpan AS CFileTimeSpan)
   DECLARE OPERATOR -= (BYREF dateSpan AS CFileTimeSpan)
   DECLARE FUNCTION GetTime () AS ULONGLONG
   DECLARE SUB SetTime (BYVAL nTime AS ULONGLONG)
   DECLARE FUNCTION GetCurrentTime () AS CFileTime
   DECLARE FUNCTION UTCToLocal () AS CFileTime
   DECLARE FUNCTION LocalToUTC () AS CFileTime
   DECLARE FUNCTION Day () AS ULONGLONG
   DECLARE FUNCTION Hour () AS ULONGLONG
   DECLARE FUNCTION Minute () AS ULONGLONG
   DECLARE FUNCTION Second () AS ULONGLONG
   DECLARE FUNCTION MilliSecond () AS ULONGLONG
   DECLARE FUNCTION Week () AS ULONGLONG
   DECLARE FUNCTION GetAsFileTime () AS FILETIME
   DECLARE FUNCTION GetAsSystemTime () AS SYSTEMTIME
   DECLARE FUNCTION Format (BYREF wszFmt AS WSTRING) AS CWSTR

END TYPE
' ========================================================================================

' ========================================================================================
' CFileTime constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CFileTime
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Constructs a CFileTime object from a date and time expressed as a 64-bit value.
' ========================================================================================
PRIVATE CONSTRUCTOR CFileTime (BYVAL nTime AS ULONGLONG)
	dwLowDateTime = CULNG(nTime)
	dwHighDateTime = CULNG(nTime shr 32)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Constructs a CFileTime object from a FILETIME structure.
' ========================================================================================
PRIVATE CONSTRUCTOR CFileTime (BYREF ft AS FILETIME)
	dwLowDateTime = ft.dwLowDateTime
	dwHighDateTime = ft.dwHighDateTime
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Constructs a CFileTime object from a SYSTEMTIME structure.
' ========================================================================================
PRIVATE CONSTRUCTOR CFileTime (BYREF st AS SYSTEMTIME)
   DIM ft AS FILETIME
   SystemTimeToFileTime @st, @ft
	dwLowDateTime = ft.dwLowDateTime
	dwHighDateTime = ft.dwHighDateTime
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CFileTime destructor
' ========================================================================================
PRIVATE DESTRUCTOR CFileTime
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the underlying m_time value from this CFileTime object.
' Example:
'   DIM cft AS CFileTime = CFileTime().GetCurrentTime()
'   DIM nTime AS LONGLONG = cft
'   print nTime
' ========================================================================================
PRIVATE OPERATOR CFileTime.CAST () AS LONGLONG
   RETURN (CULNGINT(dwHighDateTime) shl 32) OR dwLowDateTime
END OPERATOR
' ========================================================================================

' ========================================================================================
' Assigns a time value to this CFileTime object.
' nTime: A date and time expressed as a 64-bit value.
' ========================================================================================
PRIVATE OPERATOR CFileTime.LET (BYVAL nTime AS ULONGLONG)
	dwLowDateTime = CULNG(nTime)
	dwHighDateTime = CULNG(nTime shr 32)
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a FILETIME value to this CFileTime object.
' ========================================================================================
PRIVATE OPERATOR CFileTime.LET (BYREF ft AS FILETIME)
	dwLowDateTime = ft.dwLowDateTime
	dwHighDateTime = ft.dwHighDateTime
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assigns a SYSTEMTIME value to this CFileTime object.
' ========================================================================================
PRIVATE OPERATOR CFileTime.LET (BYREF st AS SYSTEMTIME)
   DIM ft AS FILETIME
   SystemTimeToFileTime @st, @ft
	dwLowDateTime = ft.dwLowDateTime
	dwHighDateTime = ft.dwHighDateTime
END OPERATOR
' ========================================================================================

' ========================================================================================
' Retrieve the time from the CFileTime object.
' Example:
'   DIM cft AS CFileTime = CFileTime().GetCurrentTime()
'   print cft.GetTime
' ========================================================================================
PRIVATE FUNCTION CFileTime.GetTime () AS ULONGLONG
   RETURN (CULNGINT(dwHighDateTime) shl 32) OR dwLowDateTime
END FUNCTION
' ========================================================================================

' ========================================================================================
' Set the date and time stored in the CFileTime object.
' ========================================================================================
PRIVATE SUB CFileTime.SetTime (BYVAL nTime AS ULONGLONG)
	dwLowDateTime = CULNG(nTime)
	dwHighDateTime = CULNG(nTime shr 32)
END SUB
' ========================================================================================

' ========================================================================================
' Adds a CFileTimeSpan value to this CFileTime object.
' ========================================================================================
PRIVATE OPERATOR CFileTime.+= (BYREF dateSpan AS CFileTimeSpan)
	this.SetTime(this.GetTime()+dateSpan.GetTimeSpan())
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CFileTimeSpan value from this CFileTime object.
' ========================================================================================
PRIVATE OPERATOR CFileTime.-= (BYREF dateSpan AS CFileTimeSpan)
	this.SetTime(this.GetTime()-dateSpan.GetTimeSpan())
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a CFileTime object that represents the current system date and time.
' ========================================================================================
PRIVATE FUNCTION CFileTime.GetCurrentTime () AS CFileTime
	DIM ft AS FILETIME
	GetSystemTimeAsFileTime(@ft)
	RETURN CFileTime(ft)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts time based on the Coordinated Universal Time (UTC) to local file time.
' ========================================================================================
PRIVATE FUNCTION CFileTime.UTCToLocal () AS CFileTime
	DIM ft AS FILETIME, ftLocal AS FILETIME
   ft = TYPE<FILETIME>(dwLowDateTime, dwHighDateTime)
	FileTimeToLocalFileTime(@ft, @ftLocal)
	RETURN CFileTime(ftLocal)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a local file time to a file time based on the Coordinated Universal Time (UTC).
' ========================================================================================
PRIVATE FUNCTION CFileTime.LocalToUTC () AS CFileTime
	DIM ft AS FILETIME, ftUTC AS FILETIME
   ft = TYPE<FILETIME>(dwLowDateTime, dwHighDateTime)
	LocalFileTimeToFileTime(@ft, @ftUTC)
	RETURN CFileTime(ftUTC)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The number of 100-nanosecond intervals that make up one day.
' ========================================================================================
PRIVATE FUNCTION CFileTime.Day () AS ULONGLONG
   RETURN CFileTime_Day
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one hour.
' ========================================================================================
PRIVATE FUNCTION CFileTime.Hour () AS ULONGLONG
   RETURN CFileTime_Hour
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one minute.
' ========================================================================================
PRIVATE FUNCTION CFileTime.Minute () AS ULONGLONG
   RETURN CFileTime_Minute
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one second.
' ========================================================================================
PRIVATE FUNCTION CFileTime.Second () AS ULONGLONG
   RETURN CFileTime_Second
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one millisecond.
' ========================================================================================
PRIVATE FUNCTION CFileTime.MilliSecond () AS ULONGLONG
   RETURN CFileTime_Millisecond
END FUNCTION
' ========================================================================================
' ========================================================================================
' The number of 100-nanosecond intervals that make up one week.
' ========================================================================================
PRIVATE FUNCTION CFileTime.Week () AS ULONGLONG
   RETURN CFileTime_Week
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the time as a FILETIME structure.
' ========================================================================================
PRIVATE FUNCTION CFileTime.GetAsFileTime () AS FILETIME
   DIM ft AS FILETIME
	ft.dwLowDateTime = dwLowDateTime
	ft.dwHighDateTime = dwHighDateTime
   RETURN ft
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the time as a SYSTEMTIME structure.
' ========================================================================================
PRIVATE FUNCTION CFileTime.GetAsSystemTime () AS SYSTEMTIME
   DIM ft AS FILETIME
	ft.dwLowDateTime = dwLowDateTime
	ft.dwHighDateTime = dwHighDateTime
   DIM st AS SYSTEMTIME
   FileTimeToSystemTime(@ft, @st)
   RETURN st
END FUNCTION
' ========================================================================================

' ========================================================================================
' Converts a CFileTime object to a string.
' This form formats the value by using the format string which contains special formatting
' codes that are preceded by a percent sign (%), as in printf.
' For more information about the formatting codes, see strftime, wcsftime in the Run-Time
' Library Reference: https://msdn.microsoft.com/en-us/library/38wh24td.aspx#coledatetime__format
' Example: print ct.Format("%A, %B %d, %Y %H:%M:%S")
' Example: print ct.FormatGmt("%A, %B %d, %Y %H:%M:%S")
' ========================================================================================
PRIVATE FUNCTION CFileTime.Format (BYREF wszFmt AS WSTRING) AS CWSTR
   DIM cwsOut AS CWSTR
   IF LEN(wszFmt) = 0 THEN RETURN cwsOut
   cwsOut = STRING(128, CHR(0))
   DIM st AS SYSTEMTIME = this.GetAsSystemTime
   DIM tmTemp AS tm
   tmTemp.tm_wday = st.wDayOfWeek
   tmTemp.tm_min = st.wMinute
   tmTemp.tm_sec = st.wSecond
   tmTemp.tm_mon = st.wMonth - 1
   tmTemp.tm_mday = st.wDay
   tmTemp.tm_hour = st.wHour
   tmTemp.tm_year = st.wYear - 1900
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
' Adds CFileTimeSpan values.
' ========================================================================================
PRIVATE OPERATOR + (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS CFileTimeSpan
   OPERATOR = CFileTimeSpan(cSpan1.m_nSpan + cSpan2.m_nSpan )
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts CFileTimeSpan values.
' ========================================================================================
PRIVATE OPERATOR - (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS CFileTimeSpan
   OPERATOR = CFileTimeSpan(cSpan1.m_nSpan - cSpan2.m_nSpan )
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators
' ========================================================================================
PRIVATE OPERATOR = (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_nSpan = cSpan2.m_nSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_nSpan <> cSpan2.m_nSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_nSpan < cSpan2.m_nSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_nSpan > cSpan2.m_nSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_nSpan <= cSpan2.m_nSpan
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
	RETURN cSpan1.m_nSpan >= cSpan2.m_nSpan
END OPERATOR
' ========================================================================================

' ========================================================================================
' Adds a CFileTimeSpan value to a CFileTime value
' ========================================================================================
PRIVATE OPERATOR + (BYREF dt AS CFileTime, BYREF dateSpan AS CFileTimeSpan) AS CFileTime
   RETURN CFileTime(dt.GetTime() + dateSpan.GetTimeSpan())
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CFileTimeSpan value from a CFileTime value
' ========================================================================================
PRIVATE OPERATOR - (BYREF dt AS CFileTime, BYREF dateSpan AS CFileTimeSpan) AS CFileTime
   RETURN CFileTime(dt.GetTime() - dateSpan.GetTimeSpan())
END OPERATOR
' ========================================================================================

' ========================================================================================
' Subtracts a CFileTime value from another CFileTime value
' ========================================================================================
PRIVATE OPERATOR - (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS CFileTimeSpan
   RETURN CFileTimeSpan(dt1.GetTime() - dt2.GetTime())
END OPERATOR
' ========================================================================================

' ========================================================================================
' Comparison operators
' ========================================================================================
PRIVATE OPERATOR = (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
   RETURN dt1.GetTime() = dt2.GetTime()
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <> (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
   RETURN dt1.GetTime() <> dt2.GetTime()
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR < (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
   RETURN dt1.GetTime() < dt2.GetTime()
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR > (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
   RETURN dt1.GetTime() > dt2.GetTime()
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR <= (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
   RETURN dt1.GetTime() <= dt2.GetTime()
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR >= (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
   RETURN dt1.GetTime() >= dt2.GetTime()
END OPERATOR
' ========================================================================================
