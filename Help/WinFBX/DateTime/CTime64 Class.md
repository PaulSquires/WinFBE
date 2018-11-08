# CTime64 Class

Represents an absolute time and date.

* **CTime64** values are based on coordinated universal time (UTC), which is equivalent to Coordinated Universal time (Greenwich Mean Time, GMT).
* A companion class, **CTimeSpan**, represents a time interval.
* The upper date limit is 12/31/3000. The lower limit is 1/1/1970 12:00:00 AM GMT.

**Include file**: CTime64.inc

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors1) | Create new **CTime64** objects initialized to the specified value. |
| [CAST Operator](#CastOp1) | Returns the **CTime64** value as a long integer. |
| [LET Operator](#LetOp1) | Assigns a value to a **CTime64** object. |
| [Operators](#Operators1) | Adds, subtracts or compares **CTime64** objects. |
| [Format](#Format) | Converts a **CTime64** object to a string. |
| [FormatGmt](#Format) | Converts a **CTime64** object to a string. |
| [GetAsFileTime](#GetAsFileTime) | Returns the time as a **FILETIME** structure. |
| [GetAsSystemTime](#GetAsSystemTime) | Returns the time as a **SYSTEMTIME** structure. |
| [GetCurrentTime](#GetCurrentTime) | Returns a **CTime64** object that represents the current system date and time. |
| [GetDay](#GetDay) | Returns the day represented by the **CTime64** object. |
| [GetDayOfWeek](#GetDayOfWeek) | Returns the day of the week represented by the **CTime64** object. |
| [GetGmtTime](#GetGmtTime) | Gets a **tm** structure that contains a decomposition of the time contained in this **CTime64** object. |
| [GetHour](#GetHour) | Returns the hour represented by the **CTime64** object. |
| [GetLocalTime](#GetLocalTime) | Gets a **tm** structure that contains a decomposition of the time contained in this **CTime64** object. |
| [GetMinute](#GetMinute) | Returns the minute represented by the **CTime64** object. |
| [GetMonth](#GetMonth) | Returns the month represented by the **CTime64** object. |
| [GetSecond](#GetSecond) | Returns the second represented by the **CTime64** object. |
| [GetTime](#GetTime) | Returns a **\_\_time64_t** (LONGLONG) value for the given **CTime64** object. |
| [GetYear](#GetYear) | Returns the year represented by the **CTime64** object. |
| [SetDateTime](#SetDateTime) | Sets the date and time of this **CTime64** object. |

# CTimeSpan Class

An amount of time, which is internally stored as the number of seconds in the time span.

* The **CTimeSpan** methods convert seconds to various combinations of days, hours, minutes, and seconds.
* The **CTimeSpan** object is stored in a \_\_time64_t structure, which is 8 bytes.
* A companion class, **CTime64**, represents an absolute time.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors2) | Create new **CTimeSpan** objects initialized to the specified value. |
| [CAST Operator](#CastOp2) | Returns the **CTimeSpan** value as a long integer. |
| [LET Operator](#LetOp2) | Assigns a value to a **CTimeSpan** object. |
| [Operators](#Operators2) | Adds, subtracts or compares **CTimeSpan** objects. |
| [GetDays](#GetDays) | Returns a value that represents the number of days in this **CTimeSpan**. |
| [GetHours](#GetHours) | Returns a value that represents the number of hours in the current day (–23 through 23). |
| [GetMinutes](#GetMinutes) | Returns a value that represents the number of minutes in the current hour (–59 through 59). |
| [GetSeconds](#GetSeconds) | Returns a value that represents the number of seconds in the current hour (–59 through 59). |
| [GetTimeSpan](#GetTimeSpan) | Returns the underlying time value from this **CTimeSpan** object. |
| [GetTotalHours](#GetTotalHours) | Returns a value that represents the total number of complete hours in this **CTimeSpan**. |
| [GetTotalMinutes](#GetTotalMinutes) | Returns a value that represents the total number of complete minutes in this **CTimeSpan**. |
| [GetTotalSeconds](#GetTotalSeconds) | Retrieves this date/time-span value expressed in seconds. |
| [SetTimeSpan](#SetTimeSpan) | Sets the value of this date/time-span value. |

# <a name="Constructors1"></a>Constructors (CTime64)

Create new **CTime64** objects initialized to the specified value.

```
CONSTRUCTOR CTime64
CONSTRUCTOR CTime64 (BYVAL timeSrc AS LONGLONG)
CONSTRUCTOR CTime64 (BYREF systimeSrc AS SYSTEMTIME)
CONSTRUCTOR CTime64 (BYREF filetimeSrc AS FILETIME)
CONSTRUCTOR CTime64 (BYVAL nYear AS LONG, BYVAL nMonth AS LONG, BYVAL nDay AS LONG, _
   BYVAL nHour AS LONG, BYVAL nMin AS LONG, BYVAL nSec AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *timeSrc* | A \_\_time64_t (LONGLONG) value. |
| *systimeSrc* | A **SYSTEMTIME** structure to be converted to a **\_\_time64_t** (LONGLONG) value and copied into the new **CTime64** object. |
| *filetimeSrc* | A **FILETIME** structure to be converted to a **\_\_time64_t** (LONGLONG) value and copied into  the new **CTime64** object. Note that **FILETIME** uses Universal Coordinated Time (UTC), so if you pass a local time in the structure, your results will be incorrect. |
| *nYear / nMonth / nDay / nHour / nMin / nSec* | Indicates the date and time values to be copied into the new **CTime64** object. |

#### Examples

Initializes a **CTime64** object from an existing **\_\_time64_t** value.

```
DIM ct1 AS CTime64 = CTime64().GetCurrentTime()
DIM ct2 AS CTime64 = ct1.GetTime
```

**Note**: You can also use DIM ct2 AS CTime64 = ct1

Initializes a **CTime64** object from individual date/time values.

```
' // Year = 2017, Month = 10 (October), Day = 9, Hour = 11, Minutes = 32, Seconds = 45
DIM ct AS CTime64 = CTime64(2017, 10, 9, 11, 32, 45)
```

Initializes a **CTime64** object from a **SYSTEMTIME** structure.

```
' // Year = 2017, Month = 10 (October), Day = 9, Hour = 11, Minutes = 32, Seconds = 45
DIM st AS SYSTEMTIME 
st.wYear = 2017
st.wMonth = 10
st.wDayOfWeek = 0
st.wDay = 9
st.wHour = 11
st.wMinute = 32
st.wSecond = 45
st.wMilliseconds = 0
' --or--
DIM st AS SYSTEMTIME = (2017, 10, 0, 9, 11, 32, 45, 0)
DIM ct AS CTime64 = st
' --or--
DIM ct AS CTime64
ct.SetDateTime(2017, 10, 9, 11, 32, 45)
DIM st AS SYSTEMTIME = ct.GetAsSystemtime
DIM ct2 AS CTime64 = st
' --or--
DIM ct AS CTime64
ct.SetDateTime(2017, 10, 9, 11, 32, 45)
DIM ct2 AS CTime64 = ct.GetAsSystemtime
```

# <a name="CastOp1"></a>CAST Operator (CTime64)

Returns the underlying value from this **CTime64** object.

```
OPERATOR CAST () AS LONGLONG
```

# <a name="LetOp1"></a>LET Operator (=) (CTime64)

Assigns a value to a **CTime64** object.

```
OPERATOR LET (BYVAL timeSrc AS LONGLONG)
OPERATOR LET (BYREF st AS SYSTEMTIME)
OPERATOR LET (BYREF ft AS FILETIME)
```

| Parameter  | Description |
| ---------- | ----------- |
| *timeSrc* | A \_\_time64_t (LONGLONG) value. |
| *st* | A **SYSTEMTIME** structure to be converted to a **\_\_time64_t** (LONGLONG) value and copied into the new **CTime64** object. |
| *ft* | A **FILETIME** structure to be converted to a **\_\_time64_t** (LONGLONG) value and copied into  the new **CTime64** object. Note that **FILETIME** uses Universal Coordinated Time (UTC), so if you pass a local time in the structure, your results will be incorrect. |

# <a name="Operators1"></a>Operators (CTime64)

Adds, subtracts or compares **CTime64** objects.

```
OPERATOR + (BYREF dt AS CTime64, BYREF dateSpan AS CTimeSpan) AS CTime64
OPERATOR - (BYREF dt AS CTime64, BYREF dateSpan AS CTimeSpan) AS CTime64
OPERATOR - (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS CTimeSpan
OPERATOR += (BYREF dateSpan AS CTimeSpan)
OPERATOR -= (BYREF dateSpan AS CTimeSpan)
OPERATOR = (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
OPERATOR <> (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
OPERATOR < (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
OPERATOR > (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
OPERATOR <= (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
OPERATOR >= (BYREF dt1 AS CTime64, BYREF dt2 AS CTime64) AS BOOLEAN
```

#### Remarks

```
+ Adds a CTimeSpan object to a CTime64 object.
- Subtracts a CTimeSpan object from a CTime64 object.
- Subtracts a CTime64 object from another CTime64 object.
+= Adds a CTimeSpan object to this CTime64 object.
-= Subtracts a CTimeSpan object from this CTime64 object.
```

# <a name="Format"></a>Format / FormatGmt (CTime64)

Converts a **CTime64** object to a string.

```
FUNCTION Format (BYREF wszFmt AS WSTRING) AS CWSTR
FUNCTION FormatGmt (BYREF wszFmt AS WSTRING) AS CWSTR
```
Formatting codes:

| Code       | Meaning     |
| ---------- | ----------- |
| %a | Abbreviated weekday name |
| %A | Full weekday name |
| %b | Abbreviated month name |
| %B | Full month name |
| %c | Date and time representation appropriate for locale |
| %d | Day of month as decimal number (01 – 31) |
| %H | Hour in 24-hour format (00 – 23) |
| %I | Hour in 12-hour format (01 – 12) |
| %j | Day of year as decimal number (001 – 366) |
| %m | Month as decimal number (01 – 12) |
| %M | Minute as decimal number (00 – 59) |
| %p | Current locale's A.M./P.M. indicator for 12-hour clock |
| %S | Second as decimal number (00 – 59) |
| %U | Week of year as decimal number, with Sunday as first day of week (00 – 53) |
| %w | Weekday as decimal number (0 – 6; Sunday is 0) |
| %W | Week of year as decimal number, with Monday as first day of week (00 – 53) |
| %x | Date representation for current locale |
| %X | Time representation for current locale |
| %y | Year without century, as decimal number (00 – 99) |
| %Y | Year with century, as decimal number |
| %z, %Z | Either the time-zone name or time zone abbreviation, depending on registry settings; no characters if time zone is unknown |
| %% | Percent sign |

The # flag may prefix any formatting code. In that case, the meaning of the format code is changed as follows.

* %#a, %#A, %#b, %#B, %#p, %#X, %#z, %#Z, %#%: # flag is ignored.
* %#c: Long date and time representation, appropriate for current locale. For example: "Tuesday, March 14, 1995, 12:41:29".
* %#x* Long date representation, appropriate to current locale. For example: "Tuesday, March 14, 1995".
* %#d, %#H, %#I, %#j, %#m, %#M, %#S, %#U, %#w, %#W, %#y, %#Y: Remove leading zeros (if any).

#### Remarks

Formats the value by using the format string which contains special formatting codes that are preceded by a percent sign (%).

**Format** uses the local time and **FormatGmt** the **Gmt** time.

#### Examples

```
print ct.Format("%A, %B %d, %Y %H:%M:%S")
print ct.FormatGmt("%A, %B %d, %Y %H:%M:%S")
```

# <a name="GetAsFileTime"></a>GetAsFileTime (CTime64)

Converts the time information stored in the **CTime64** object to a Win32–compatible FILETIME structure. A **FILETIME** structure contains a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (UTC).

```
FUNCTION GetAsFileTime () AS FILETIME
```

# <a name="GetAsSystemTime"></a>GetAsSystemTime (CTime64)

Converts the time information stored in the **CTime64** object to a Win32–compatible **SYSTEMTIME** structure.

```
FUNCTION GetAsSystemTime () AS SYSTEMTIME
```

# <a name="GetCurrentTime"></a>GetCurrentTime (CTime64)

Returns a **CTime64** object that represents the current time.

```
FUNCTION GetCurrentTime () AS CTime64
```

# <a name="GetDay"></a>GetDay (CTime64)

Returns the day represented by the **CTime64** object.

```
FUNCTION GetDay () AS LONG
```

#### Example

```
DIM ct AS CTime64
ct = ct.GetTime
PRINT ct.GetDay
```

# <a name="GetDayOfWeek"></a>GetDayOfWeek (CTime64)

Returns the day of the week represented by the **CTime64** object.

```
FUNCTION GetDayOfWeek () AS LONG
```

#### Example

```
DIM ct AS CTime64
ct = ct.GetTime
PRINT ct.GetDayOfWeek
```

# <a name="GetGmtTime"></a>GetGmtTime (CTime64)

Gets a **tm** structure that contains a decomposition of the time contained in this **CTime64** object.

```
FUNCTION GetGmtTime () AS tm
```

#### Remarks

The fields of the structure type tm store the following values, each of which is an int:

tm_sec : Seconds after minute (0 – 59).<br>
tm_min : Minutes after hour (0 – 59).<br>
tm_hour : Hours after midnight (0 – 23).<br>
tm_mday : Day of month (1 – 31).<br>
tm_mon : Month (0 – 11; January = 0).<br>
tm_year : Year (current year minus 1900).<br>
tm_wday : Day of week (0 – 6; Sunday = 0).<br>
tm_yday : Day of year (0 – 365; January 1 = 0).<br>
tm_isdst : Positive value if daylight saving time is in effect; 0 if daylight saving time is not in effect; negative value if status of daylight saving time is unknown.

# <a name="GetHour"></a>GetHour (CTime64)

Returns the hour represented by the **CTime64** object.

```
FUNCTION GetHour () AS LONG
```

#### Example

```
DIM ct AS CTime64
ct = ct.GetTime
PRINT ct.GetHour
```

# <a name="GetLocalTime"></a>GetLocalTime (CTime64)

Gets a **tm** structure that contains a decomposition of the time contained in this **CTime64** object.

```
FUNCTION GetLocalTime () AS tm
```

#### Remarks

The fields of the structure type tm store the following values, each of which is an int:

tm_sec : Seconds after minute (0 – 59).<br>
tm_min : Minutes after hour (0 – 59).<br>
tm_hour : Hours after midnight (0 – 23).<br>
tm_mday : Day of month (1 – 31).<br>
tm_mon : Month (0 – 11; January = 0).<br>
tm_year : Year (current year minus 1900).<br>
tm_wday : Day of week (0 – 6; Sunday = 0).<br>
tm_yday : Day of year (0 – 365; January 1 = 0).<br>
tm_isdst : Positive value if daylight saving time is in effect; 0 if daylight saving time is not in effect; negative value if status of daylight saving time is unknown.

# <a name="GetMinute"></a>GetMinute (CTime64)

Returns the minute represented by the **CTime64** object.

```
FUNCTION GetMinute () AS LONG
```

#### Example

```
DIM ct AS CTime64
ct = ct.GetTime
PRINT ct.GetMinute
```

# <a name="GetMonth"></a>GetMonth (CTime64)

Returns the month represented by the **CTime64** object.

```
FUNCTION GetMonth () AS LONG
```

#### Example

```
DIM ct AS CTime64
ct = ct.GetTime
PRINT ct.GetMonth
```

# <a name="GetSecond"></a>GetSecond (CTime64)

Returns the second represented by the **CTime64** object.

```
FUNCTION GetSecond () AS LONG
```

#### Example

```
DIM ct AS CTime64
ct = ct.GetTime
PRINT ct.GetSecond
```

# <a name="GetTime"></a>GetTime (CTime64)

Returns a **\_\_time64_t** (LONGLONG) value for the given **CTime64** object.

```
FUNCTION GetTime () AS LONGLONG
```

#### Example

```
DIM ct AS CTime64
ct = ct.GetTime
```

# <a name="GetYear"></a>GetYear (CTime64)

Returns the year represented by the **CTime64** object.

```
FUNCTION GetYear () AS LONG
```

#### Example

```
DIM ct AS CTime64
ct = ct.GetTime
PRINT ct.GetYear
```

# <a name="SetDateTime"></a>SetDateTime (CTime64)

Sets the date and time of this **CTime64** object.

```
FUNCTION SetDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, _
   BYVAL nHour AS WORD = 0, BYVAL nMin AS WORD = 0, BYVAL nSec AS WORD = 0) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | The year (1970-3000). |
| *nMonth* | The month (1-12). |
| *nDay* | The day (1-31). |
| *nHour* | The hour (0-23). |
| *nMinute* | The minutes (0-59). |
| *nHour* | The seconds (0-59). |

#### Example

```
' // Year = 2017, Month = 10 (October), Day = 9, Hour = 11, Minutes = 32, Seconds = 45
DIM ct AS CTime64 = CTime64(2017, 10, 9, 11, 32, 45)
```

# <a name="Constructors2"></a>Constructors (CTimeSpan)

Create new **CTimespan** objects initialized to the specified value.

```
CONSTRUCTOR CTimeSpan
CONSTRUCTOR CTimeSpan (BYVAL lltime AS LONGLONG)
CONSTRUCTOR CTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lltime* | A **\_\_time64_t** (LONGLONG) value. |
| *lDays / nHours / nMins / nSecs* | Indicates the day and time values to be copied into the new **CTimeSpan** object. |

# <a name="CastOp2"></a>CAST Operator (CTimeSpan)

Returns the underlying value from this **CTimeSpan** object.

```
OPERATOR CAST () AS LONGLONG
```

# <a name="LetOp2"></a>LET Operator (=) (CTimeSpan)

Assigns a value to a **CTimeSpan** object.

```
OPERATOR LET (BYVAL nSpan AS LONGLONG)
OPERATOR LET (BYREF cSpan AS CTimeSpan)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nSpan* | A \_\_time64_t (LONGLONG) value. |
| *cSpan* | A **CTimeSpan** object. |

# <a name="Operators2"></a>Operators (CTimeSpan)

Adds, subtracts or compares **CTimeSpan** objects.

```
OPERATOR + (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS CTime64
OPERATOR - (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS CTime64
OPERATOR - (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS CTimeSpan
OPERATOR += (BYREF cSpan AS CTimeSpan)
OPERATOR -= (BYREF cSpan AS CTimeSpan)
OPERATOR = (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
OPERATOR <> (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
OPERATOR < (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
OPERATOR > (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
OPERATOR <= (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
OPERATOR >= (BYREF cSpan1 AS CTimeSpan, BYREF cSpan2 AS CTimeSpan) AS BOOLEAN
```

#### Remarks

```
+ Adds a CTimeSpan value to another CTimeSpan object.
- Subtracts a CTimeSpan value from another CTimeSpan object.
+= Adds a CTimeSpan value to this CTimeSpan object.
-= Subtracts a CTimeSpan value from this CTimeSpan object.
```

# <a name="GetDays"></a>GetDays (CTimeSpan)

Returns a value that represents the number of days in this **CTimeSpan**.

```
FUNCTION GetDays () AS LONG
```

# <a name="GetHours"></a>GetHours (CTimeSpan)

Returns a value that represents the number of hours in the current day (–23 through 23).

```
FUNCTION GetHours () AS LONG
```

# <a name="GetMinutes"></a>GetMinutes (CTimeSpan)

Returns a value that represents the number of minutes in the current hour (–59 through 59).

```
FUNCTION GetMinutes () AS LONG
```

# <a name="GetSeconds"></a>GetSeconds (CTimeSpan)

Returns a value that represents the number of seconds in the current hour (–59 through 59).

```
FUNCTION GetSeconds () AS LONG
```

# <a name="GetTimeSpan"></a>GetTimeSpan (CTimeSpan)

Returns the underlying time value from this **CTimeSpan** object.

```
FUNCTION GetTimespan () AS LONG
```

# <a name="GetTotalHours"></a>GetTotalHours (CTimeSpan)

Returns a value that represents the total number of complete hours in this **CTimeSpan**.

```
FUNCTION GetTotalHours () AS LONG
```

# <a name="GetTotalMinutes"></a>GetTotalMinutes (CTimeSpan)

Returns a value that represents the total number of complete minutes in this **CTimeSpan**.

```
FUNCTION GetTotalMinutes () AS LONG
```

# <a name="GetTotalSeconds"></a>GetTotalSeconds (CTimeSpan)

Retrieves this date/time-span value expressed in seconds.

```
FUNCTION GetTotalSeconds () AS LONG
```

# <a name="SetTimeSpan"></a>SetTimeSpan (CTimeSpan)

Sets the value of this date/time-span value.

```
SUB SetTimeSpan (BYVAL lltime AS LONGLONG)
SUB SetTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lltime* | A \_\_time64_t (LONGLONG) value. |
| *lDays / nHours / nMins / nSecs* | Indicates the day and time values to be copied into the **CTimeSpan** object. |
