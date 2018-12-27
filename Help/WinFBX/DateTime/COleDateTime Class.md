# COleDateTime Class

Encapsulates the **DATE** data type that is used in OLE automation.

It is one of the possible types for the VARIANT data type of OLE automation. A **COleDateTime** value represents an absolute date and time value. The **DATE** type is implemented as a floating-point value. Days are measured from December 30, 1899, at midnight. The **COleDateTime** class handles dates from January 1, 100, through December 31, 9999. The **COleDateTime** class uses the Gregorian calendar; it does not support Julian dates. **COleDateTime** ignores Daylight Saving Time. This type is also used to represent date-only or time-only values. By convention, the date 0 (December 30, 1899) is used for time-only values and the time 00:00 (midnight) is used for date-only values.

**Caution**: Note that although day values become negative before midnight on December 30, 1899, time-of-day values do not. For example, 6:00 AM is always represented by a fractional value 0.25 regardless of whether the integer representing the day is positive (after December 30, 1899) or negative (before December 30, 1899). This means that a simple floating point comparison would erroneously sort a **COleDateTime** representing 6:00 AM on 12/29/1899 as later than one representing 7:00 AM on the same day.

**Include file**: COleDateTime.inc

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors1) | Create new **COleDateTime** objects initialized to the specified value. |
| [CAST Operator](#CastOp1) | Returns the **COleDateTime** value as a long integer. |
| [LET Operator](#LetOp1) | Assigns a value to a **COleDateTime** object. |
| [Operators](#Operators1) | Adds, subtracts or compares **COleDateTime** objects. |
| [Format](#Format) | Converts a **COleDateTime** object to a string. |
| [Format (Overload)](#Format2) | Converts a **COleDateTime** object to a string. |
| [GetAsDBTIMESTAMP](#GetAsDBTIMESTAMP) | Returns the date/time the of this **COleDateTime** object as a **DBTIMESTAMP** data structure. |
| [GetAsSystemTime](#GetAsSystemTime) | Returns the date/time the of this **COleDateTime** object as a **SYSTEMTIME** data structure. |
| [GetAsUdate](#GetAsUdate) | Returns the date/time the of this **COleDateTime** object as a UDATE structure. |
| [GetCurrentTime](#GetCurrentTime) | Returns the current date/time value in local time. |
| [GetDay](#GetDay) | Gets the day represented by this date/time value. |
| [GetDayOfWeek](#GetDayOfWeek) | Gets the day of the week represented by this date/time value. |
| [GetDayOfYear](#GetDayOfYear) | Gets the day of the year represented by this date/time value. |
| [GetHour](#GetHour) | Gets the hour represented by this date/time value. |
| [GetLocalTime](#GetLocalTime) | Returns the current date/time value in local time. |
| [GetMinute](#GetMinute) | Gets the minute represented by this date/time value. |
| [GetMonth](#GetMonth) | Gets the month represented by this date/time value. |
| [GetSecond](#GetSecond) | Gets the second represented by this date/time value. |
| [GetStatus](#GetStatus) | Gets the status (validity) of a given **COleDateTime** object. |
| [GetSystemTime](#GetSystemTime) | Returns the current date/time value in Coordinated Universal Time (UTC). |
| [GetYear](#GetYear) | Gets the year represented by this date/time value. |
| [SetDate](#SetDate) | Sets the date of this **COleDateTime** object. |
| [SetDateTime](#SetDateTime) | Sets the date and time of this **COleDateTime** object. |
| [SetStatus](#SetStatus) | Sets the status (validity) of a given **COleDateTime** object. |
| [SetTime](#SetTime) | Sets the time of this **COleDateTime** object. |

# COleDateTimeSpan Class

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors2) | Create new **COleDateTimeSpan** objects initialized to the specified value. |
| [CAST Operator](#CastOp2) | Returns the **COleDateTimeSpan** value as a long integer. |
| [LET Operator](#LetOp2) | Assigns a value to a **COleDateTimeSpan** object. |
| [Operators](#Operators2) | Adds, subtracts or compares **COleDateTimeSpan** objects. |
| [GetDays](#GetDays) | Retrieves the day portion of this date/time-span value. |
| [GetHours](#GetHours) | Retrieves the hour portion of this date/time-span value. |
| [GetMinutes](#GetMinutes) | Retrieves the minutes portion of this date/time-span value. |
| [GetSeconds](#GetSeconds) | Retrieves the seconds portion of this date/time-span value. |
| [GetStatus](#GetStatusSpan) | Gets the status (validity) of a given **COleDateTimeSpan** object. |
| [GetTotalDays](#GetTotalDays) | Retrieves this date/time-span value expressed in days. |
| [GetTotalHours](#GetTotalHours) | Retrieves this date/time-span value expressed in hours. |
| [GetTotalMinutes](#GetTotalMinutes) | Retrieves this date/time-span value expressed in minutes. |
| [GetTotalSeconds](#GetTotalSeconds) | Retrieves this date/time-span value expressed in seconds. |
| [SetDateTimeSpan](#SetDateTimeSpan) | Sets the date and time of this **COleDateTimeSpan** object. |
| [SetStatus](#SetStatusSpan) | Sets the status (validity) of a given **COleDateTimeSpan** object. |

# <a name="Constructors1"></a>Constructors (COleDateTime)

Create new **COleDateTime** objects initialized to the specified value.

```
CONSTRUCTOR COleDateTime
CONSTRUCTOR COleDateTime (BYREF dateSrc AS COleDateTime)
CONSTRUCTOR COleDateTime (BYREF varSrc AS VARIANT)
CONSTRUCTOR COleDateTime (BYVAL dtSrc AS DATE_)
CONSTRUCTOR COleDateTime (BYVAL timeSrc AS LONGLONG)
CONSTRUCTOR COleDateTime (BYREF systimeSrc AS SYSTEMTIME)
CONSTRUCTOR COleDateTime (BYREF filetimeSrc AS FILETIME)
CONSTRUCTOR COleDateTime (BYVAL wDosDate AS DWORD, BYVAL wDosTime AS WORD)
CONSTRUCTOR COleDateTime (BYREF dbts AS DBTIMESTAMP)
CONSTRUCTOR COleDateTime (BYREF ud AS UDATE)
CONSTRUCTOR COleDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, _
   BYVAL nHour AS WORD, BYVAL nMin AS WORD, BYVAL nSec AS WORD)
CONSTRUCTOR COleDateTime (BYREF wsz AS WSTRING, BYVAL dwFlags AS DWORD = 0, _
   BYVAL lcid AS LCID = LANG_USER_DEFAULT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dateSrc* | An existing **COleDateTime** object to be copied into the new **COleDateTime** object. |
| *varSrc* | An existing VARIANT data structure to be converted to a date/time value (VT_DATE) and copied into the new **COleDateTime** object. |
| *dtSrc* | A date/time ( DATE) value to be copied into the new **COleDateTime** object. |
| *timeSrc* | A LongInt value to be converted to a date/time value and copied into the new **COleDateTime** object. |
| *systimeSrc* | A **SYSTEMTIME** structure to be converted to a date/time value and copied into the new **COleDateTime** object. |
| *filetimeSrc* | A **FILETIME** structure to be converted to a date/time value and copied into the new **COleDateTime** object. Note that **FILETIME** uses Universal Coordinated Time (UTC), so if you pass a local time in the structure, your results will be incorrect. See File Times in the Windows SDK for more information. |
| *nYear / nMonth / nDay / nHour / nMin / nSec* | Indicate the date and time values to be copied into the new **COleDateTime** object. |
| *wDosDate / wDosTime* | MS-DOS date and time values to be converted to a date/time value and copied into the new **COleDateTime** object. |
| *dbts* | A reference to a **DBTimeStamp** structure containing the current local time. |
| *ud* | The **UDATE** value is converted and copied into this **COleDateTime** object. If the conversion is successful, the status of this object is set to valid; if unsuccessful, it is set to invalid. A **UDATE** structure represents an "unpacked" date. See the function **VarDateFromUdate** for more details. |
| *wsz* | A null-terminated unicode string which is to be parsed. This parameter can take a variety of formats. For example, the following strings contain acceptable date/time formats:<br>"25 January 1996"<br>"8:30:00"<br>"20:30:00"<br>"January 25, 1996 8:30:00"<br>"8:30:00 Jan. 25, 1996"<br>"1/25/1996 8:30:00" // always specify the full year, even in a 'short date' format.<br>Note that the locale ID will also affect whether the string format is acceptable for conversion to a date/time value. In the case of VAR_DATEVALUEONLY, the time value is set to time 0, or midnight. In the case of VAR_TIMEVALUEONLY, the date value is set to date 0, meaning 30 December 1899. |
| *dwFlags* | Indicates flags for locale settings and parsing.<br>LOCALE_NOUSEROVERRIDE Use the system default locale settings, rather than custom user settings.<br>VAR_TIMEVALUEONLY Ignore the date portion during parsing.<br>VAR_DATEVALUEONLY Ignore the time portion during parsing. |
| *lcid* | Indicates locale ID to use for the conversion. Default value: LANG_USER_DEFAULT. |

#### Remarks

The status of the new **COleDateTime** object is set to valid.

# <a name="CastOp1"></a>CAST Operator (COleDateTime)

Returns the underlying **DATE** value from this **COleDateTime** object.

```
OPERATOR CAST () AS DATE_
```

# <a name="LetOp1"></a>LET Operator (=) (COleDateTime)

Assigns a value to a **COleDateTime** object.

```
OPERATOR LET (BYREF dateSrc AS COleDateTime)
OPERATOR LET (BYREF varSrc AS VARIANT)
OPERATOR LET (BYVAL dtSrc AS DATE_)
OPERATOR LET (BYVAL timeSrc AS LONGLONG)
OPERATOR LET (BYREF systimeSrc AS SYSTEMTIME)
OPERATOR LET (BYREF filetimeSrc AS FILETIME)
OPERATOR LET (BYREF dbts AS DBTIMESTAMP)
OPERATOR LET (BYREF ud AS UDATE)
OPERATOR LET (BYREF wsz AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dateSrc* | An existing **COleDateTime** object to be copied into the new **COleDateTime** object. |
| *varSrc* | An existing VARIANT data structure to be converted to a date/time value (VT_DATE) and copied into the new **COleDateTime** object. |
| *dtSrc* | A date/time ( DATE) value to be copied into the new **COleDateTime** object. |
| *timeSrc* | A LongInt value to be converted to a date/time value and copied into the new **COleDateTime** object. |
| *systimeSrc* | A **SYSTEMTIME** structure to be converted to a date/time value and copied into the new **COleDateTime** object. |
| *filetimeSrc* | A **FILETIME** structure to be converted to a date/time value and copied into the new **COleDateTime** object. Note that **FILETIME** uses Universal Coordinated Time (UTC), so if you pass a local time in the structure, your results will be incorrect. See File Times in the Windows SDK for more information. |
| *dbts* | A reference to a **DBTimeStamp** structure containing the current local time. |
| *ud* | The **UDATE** value is converted and copied into this **COleDateTime** object. If the conversion is successful, the status of this object is set to valid; if unsuccessful, it is set to invalid. A **UDATE** structure represents an "unpacked" date. See the function **VarDateFromUdate** for more details. |
| *wsz* | A null-terminated unicode string which is to be parsed. This parameter can take a variety of formats. For example, the following strings contain acceptable date/time formats:<br>"25 January 1996"<br>"8:30:00"<br>"20:30:00"<br>"January 25, 1996 8:30:00"<br>"8:30:00 Jan. 25, 1996"<br>"1/25/1996 8:30:00" // always specify the full year, even in a 'short date' format.<br>Note that the locale ID will also affect whether the string format is acceptable for conversion to a date/time value. In the case of VAR_DATEVALUEONLY, the time value is set to time 0, or midnight. In the case of VAR_TIMEVALUEONLY, the date value is set to date 0, meaning 30 December 1899. |

# <a name="Operators1"></a>Operators (COleDateTime)

Adds, subtracts or compares **COleDateTime** objects.

```
OPERATOR + (BYREF dt AS COleDateTime, BYREF dateSpan AS COleDateTimeSpan) AS COleDateTime
OPERATOR - (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS COleDateTimeSpan
OPERATOR - (BYREF dt AS COleDateTime, BYREF dateSpan AS COleDateTimeSpan) AS COleDateTime
OPERATOR += (BYREF dateSpan AS COleDateTimeSpan)
OPERATOR -= (BYREF dateSpan AS COleDateTimeSpan)
OPERATOR = (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
OPERATOR <> (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
OPERATOR < (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
OPERATOR > (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
OPERATOR <= (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
OPERATOR >= (BYREF dt1 AS COleDateTime, BYREF dt2 AS COleDateTime) AS BOOLEAN
```

# <a name="Format"></a>Format (COleDateTime)

Converts a **COleDateTime** to a formatted string.

```
FUNCTION Format (BYREF wszFmt AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFmt* | A format string which contains special formatting codes that are preceded by a percent sign (%). |

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

# <a name="Format2"></a>Format (Overload) (COleDateTime)

Converts a **COleDateTime** to a formatted string.

```
FUNCTION Format (BYVAL dwFlags AS DWORD = 0, BYVAL lcid AS LCID = LANG_USER_DEFAULT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwFlags* | One or more of the following flags. |
| *lcid* | The locale identifier. Default value: LANG_USER_DEFAULT. |

| Flag       | Meaning |
| ---------- | ----------- |
| LOCALE_NOUSEROVERRIDE | Uses the system default locale settings, rather than custom locale settings. |
| VAR_CALENDAR_HIJRI | If set then the Hijri calendar is used. Otherwise the calendar set in the control panel is used. |
| VAR_CALENDAR_THAI | If set then the Buddhist year is used. |
| VAR_CALENDAR_GREGORIAN | If set the Gregorian year is used. |
| VAR_FOURDIGITYEARS | Use 4-digit years instead of 2-digit years. |
| VAR_TIMEVALUEONLY | Omits the date portion of a VT_DATE and returns only the time. Applies to conversions to or from dates. |
| VAR_DATEVALUEONLY | Omits the time portion of a VT_DATE and returns only the date. Applies to conversions to or from dates. |

#### Return value

The **COleDateTime** as a formatted string.

# <a name="GetAsDBTIMESTAMP"></a>GetAsDBTIMESTAMP (COleDateTime)

Returns the date/time the of this **COleDateTime** object as a **DBTIMESTAMP** data structure.

```
FUNCTION GetAsDBTIMESTAMP (BYREF dbts AS DBTIMESTAMP) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbts* | A reference to a **DBTIMESTAMP** structure to receive the converted date/time value from the **COleDateTime** object. |

#### Return value

TRUE or FALSE.

# <a name="GetAsSystemTime"></a>GetAsSystemTime (COleDateTime)

Returns the date/time the of this **COleDateTime** object as a **SYSTEMTIME** data structure.

```
FUNCTION GetAsSystemTime (BYREF sysTime AS SYSTEMTIME) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *sysTime* | A reference to a **SYSTEMTIME** structure to receive the converted date/time value from the **COleDateTime** object. |

#### Return value

TRUE or FALSE.

# <a name="GetAsUdate"></a>GetAsUdate (COleDateTime)

Returns the date/time the of this **COleDateTime** object as a **UDATE** structure.

```
FUNCTION GetAsUdate (BYREF ud AS UDATE) AS BOOLEAN
```

#### Return value

TRUE or FALSE.

# <a name="GetCurrentTime"></a>GetCurrentTime (COleDateTime)

Returns the current date/time value in local time.

```
FUNCTION GetCurrentTime () AS COleDateTime
```

#### Return value

The current date/time as a **COleDateTime**.

#### Example

```
DIM ct AS COleDateTime = COleDateTime().GetCurrentTime()
```
--or--
```
DIM ct AS COleDateTime
ct = ct.GetCurrentTime()
```

# <a name="GetDay"></a>GetDay (COleDateTime)

Returns the day represented by this date/time value. Valid return values range between 1 and 31.

```
FUNCTION GetDay () AS WORD
```

# <a name="GetDayOfWeek"></a>GetDayOfWeek (COleDateTime)

Returns the day of the week represented by this date/time value. Valid return values range between 1 and 7, where 1=Sunday, 2=Monday, and so on.

```
FUNCTION GetDayOfWeek () AS WORD
```

# <a name="GetDayOfYear"></a>GetDayOfYear (COleDateTime)

Returns the day of the year represented by this date/time value. Valid return values range between 1 and 366, where January 1 = 1.

```
FUNCTION GetDayOfYear () AS WORD
```

# <a name="GetHour"></a>GetHour (COleDateTime)

Returns the hour represented by this date/time value. Valid return values range between 0 and 23.

```
FUNCTION GetHour () AS WORD
```

# <a name="GetLocalTime"></a>GetLocalTime (COleDateTime)

Returns the current date/time value in local time.

```
FUNCTION GetLocalTime () AS COleDateTime
```

# <a name="GetMinute"></a>GetMinute (COleDateTime)

Returns the minute represented by this date/time value. Valid return values range between 0 and 59.

```
FUNCTION GetMinute () AS WORD
```

# <a name="GetMonth"></a>GetMonth (COleDateTime)

Returns the month represented by this date/time value. Valid return values range between 1 and 12.

```
FUNCTION GetMonth () AS WORD
```

# <a name="GetSecond"></a>GetSecond (COleDateTime)

Returns the second represented by this date/time value. Valid return values range between 0 and 59.

```
FUNCTION GetSecond () AS WORD
```

# <a name="GetStatus"></a>GetStatus (COleDateTime)

Gets the status (validity) of a given **COleDateTime** object.

```
FUNCTION GetStatus () AS BOOLEAN
```

# <a name="GetSystemTime"></a>GetSystemTime (COleDateTime)

Returns the current date/time value in Coordinated Universal Time (UTC).

```
FUNCTION GetSystemTime () AS COleDateTime
```

# <a name="GetYear"></a>GetYear (COleDateTime)

Returns the year represented by this date/time value. Valid return values range between 100 and 9999, which includes the century.

```
FUNCTION GetYear () AS WORD
```

# <a name="SetDate"></a>SetDate (COleDateTime)

Sets the date of this **COleDateTime** object.

```
FUNCTION SetDate (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear / nMonth / nDay* | Indicates the date values to be copied into the **COleDateTime** object. |

# <a name="SetDateTime"></a>SetDateTime (COleDateTime)

Sets the date and time of this **COleDateTime** object.

```
FUNCTION SetDateTime (BYVAL nYear AS WORD, BYVAL nMonth AS WORD, BYVAL nDay AS WORD, _
   BYVAL nHour AS WORD, BYVAL nMin AS WORD, BYVAL nSec AS WORD)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear / nMonth / nDay / nHour / nMin / nSec* | Indicates the date and time values to be copied into the **COleDateTime** object. |

# <a name="SetStatus"></a>SetStatus (COleDateTime)

Sets the status (validity) of a given **COleDateTime** object.

```
FUNCTION SetStatus (BYVAL nStatus AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStatus* | TRUE or FALSE. |

# <a name="SetTime"></a>SetTime (COleDateTime)

Sets the time of this **COleDateTime** object.

```
FUNCTION SetTime (BYVAL nHour AS WORD, BYVAL nMin AS WORD, BYVAL nSec AS WORD)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nHour / nMin / nSec* |  Indicates the time values to be copied into the **COleDateTime** object. |

# <a name="Constructors2"></a>Constructors (COleDateTimeSpan)

All these constructors create new **COleDateTimeSpan** objects initialized to the specified value.

```
CONSTRUCTOR COleDateTimeSpan
CONSTRUCTOR COleDateTimeSpan (dblSpanSrc AS DOUBLE)
CONSTRUCTOR COleDateTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dblSpanSrc* | The number of days to be copied into the new **COleDateTimeSpan** object. |
| *lDays / nHours / nMins / nSecs* | Indicates the day and time values to be copied into the new **COleDateTimeSpan** object. |

#### Remarks

The status of the new **COleDateTimeSpan** object is set to valid.

# <a name="CastOp2"></a>CAST Operator (COleDateTimeSpan)

Returns the underlying **DATE** value from this **COleDateTimeSpan** object.

```
OPERATOR CAST () AS DOUBLE
```

# <a name="LetOp2"></a>LET Operator (=) (COleDateTimeSpan)

Assigns a value to a **COleDateTimeSpan** object.

```
OPERATOR LET (BYVAL dblSpanSrc AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dblSpanSrc* | The number of days to be copied into the **COleDateTimeSpan** object. |

# <a name="Operators2"></a>Operators (COleDateTimeSpan)

Adds, subtracts or compares **COleDateTimeSpan** objects.

```
OPERATOR + (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS COleDateTimeSpan
OPERATOR - (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS COleDateTimeSpan
OPERATOR - (BYREF dateSpan AS COleDateTimeSpan) AS COleDateTimeSpan
OPERATOR += (BYREF dateSpan AS COleDateTimeSpan)
OPERATOR -= (BYREF dateSpan AS COleDateTimeSpan)
OPERATOR = (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS BOOLEAN
OPERATOR <> (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSPan2 AS COleDateTimeSpan) AS BOOLEAN
OPERATOR < (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS BOOLEAN
OPERATOR > (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS BOOLEAN
OPERATOR <= (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS BOOLEAN
OPERATOR >= (BYREF dateSpan1 AS COleDateTimeSpan, BYREF dateSpan2 AS COleDateTimeSpan) AS BOOLEAN
```

# <a name="GetDays"></a>GetDays (COleDateTimeSpan)

Retrieves the day portion of this date/time-span value. The return values from this function range between approximately – 3,615,000 and 3,615,000.

```
FUNCTION GetDays () AS LONGINT
```

# <a name="GetHours"></a>GetHours (COleDateTimeSpan)

Retrieves the hour portion of this date/time-span value. The return values from this function range between – 23 and 23.

```
FUNCTION GetHours () AS LONGINT
```

# <a name="GetMinutes"></a>GetMinutes (COleDateTimeSpan)

Retrieves the minutes portion of this date/time-span value. The return values from this function range between – 59 and 59.

```
FUNCTION GetMinutes () AS LONGINT
```

# <a name="GetSeconds"></a>GetSeconds (COleDateTimeSpan)

Retrieves the seconds portion of this date/time-span value. The return values from this function range between – 59 and 59.

```
FUNCTION GetSeconds () AS LONGINT
```

# <a name="GetStatusSpan"></a>GetStatus (COleDateTimeSpan)

Gets the status (validity) of a given **COleDateTimeSpan** object.

```
FUNCTION GetStatus () AS BOOLEAN
```

# <a name="GetTotalDays"></a>GetTotalDays (COleDateTimeSpan)

Retrieves this date/time-span value expressed in days.

```
FUNCTION GetTotalDays () AS LONGINT
```

# <a name="GetTotalHours"></a>GetTotalHours (COleDateTimeSpan)

Retrieves this date/time-span value expressed in hours.

```
FUNCTION GetTotalHours () AS LONGINT
```

# <a name="GetTotalMinutes"></a>GetTotalMinutes (COleDateTimeSpan)

Retrieves this date/time-span value expressed in minutes.

```
FUNCTION GetTotalMinutes () AS LONGINT
```

# <a name="GetTotalSeconds"></a>GetTotalSeconds (COleDateTimeSpan)

Retrieves this date/time-span value expressed in seconds.

```
FUNCTION GetTotalSeconds () AS LONGINT
```

# <a name="SetDateTimeSpan"></a>SetDateTimeSpan (COleDateTimeSpan)

Sets the date and time of this **COleDateTimeSpan** object.

```
FUNCTION SetDateTimeSpan (BYVAL lDays AS LONG, BYVAL nHours AS LONG, BYVAL nMins AS LONG, BYVAL nSecs AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lDays / nHours / nMins / nSecs* | Indicates the day and time values to be copied into the **COleDateTimeSpan** object. |

#### Remarks

The status of the new **COleDateTimeSpan** object is set to valid.

# <a name="SetStatusSpan"></a>SetStatus (COleDateTimeSpan)

Sets the status (validity) of a given **COleDateTimeSpan** object.

```
FUNCTION SetStatus (BYVAL nStatus AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStatus* | TRUE or FALSE. |
