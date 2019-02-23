# CPowerTime Class

A **CPowerTime** object contains a date and time value, allowing easy calculations. The date and time value is stored as a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601. A nanosecond is one-billionth of a second.

The following static const member variables are provided to simplify calculations (number of 100-nanosecond intervals):

```
CPowerTime_Millisecond    10000ull
CPowerTime_Second         CPowerTime_Millisecond * 1000
CPowerTime_Minute         CPowerTime_Second * 60
CPowerTime_Hour           CPowerTime_Minute * 60
CPowerTime_Day            CPowerTime_Hour * 24
CPowerTime_Week           CPowerTime_Day * 7
```

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors1) | Create new **CPowerTime** objects initialized to the specified value. |
| [CAST Operator](#CastOp1) | Returns the **CPowerTime** value as an unsigned long integer. |
| [LET Operator](#LetOp1) | Assigns a value to a **CPowerTime** object. |
| [Operators](#Operators1) | Adds, subtracts or compares **CPowerTime** objects. |
| [AddDays](#AddDays) | Adds the specified number of days to this **CPowerTime** object. You can subtract days by using a negative number. |
| [AddHours](#AddHours) | Adds the specified number of hours to this **CPowerTime** object. You can subtract hours by using a negative number. |
| [AddMinutes](#AddMinutes) | Adds the specified number of minutes to this **CPowerTime** object. You can subtract minutes by using a negative number. |
| [AddMonths](#AddMonths) | Adds the specified number of months to this **CPowerTime** object. You can subtract months by using a negative number. |
| [AddMSeconds](#AddMSeconds) | Adds the specified number of milliseconds to this **CPowerTime** object. You can subtract milliseconds by using a negative number. |
| [AddSeconds](#AddSeconds) | Adds the specified number of seconds to this **CPowerTime** object. You can subtract seconds by using a negative number. |
| [AddYears](#AddYears) | Adds the specified number of years to this **CPowerTime** object. You can subtract years by using a negative number. |
| [AstroDay](#AstroDay) | Returns the Astronomical Day for any given date. |
| [AstroDayOfWeek](#AstroDayOfWeek) | Returns the Astronomical Day for any given date. |
| [DateSerial](#DateSerial) | Gets/sets the date and time as a FreeBasic date serial. |
| [DateString](#DateString) | Returns the date as a string based on the specified mask, e.g. "dd-MM-yyyy". |
| [Day](#Day) | Returns the Day component of the **CPowerTime** object. It is a  value in the range of 1-31. |
| [DaysInMonth](#DaysInMonth) | Returns the number of days in the specified month. |
| [DayOfWeek](#DayOfWeek) | Returns the Day-of-Week component of the **CPowerTime** object. |
| [DayOfYear](#DayOfYear) | Returns the day of the year, where Jan 1 is the first day of the year. |
| [DayOfWeekString](#DayOfWeekstring) | Returns the Day-of-Week component of the **CPowerTime** object as a string. |
| [DaysDiff](#DaysDiff) | Returns the days of difference between two dates. |
| [DaysInYear](#DaysInYear) | Returns the number of days of the year. |
| [Format](#Format) | Converts a **CPowerTime** object to a string. |
| [GetAsFileTime](#GetAsFileTime) | Returns the date and time as a **FILETIME** structure. |
| [GetAsJulianDate](#GetAsJulianDate) | Returns the date as a Julian date. |
| [GetAsSystemTime](#GetAsSystemTime) | Returns the date and time as a **SYSTEMTIME** structure. |
| [GetCurrentTime](#GetCurrentTime) | Returns a **CPowerTime** object that represents the current system date and time. |
| [GetFileTime](#GetFileTime) | Returns the value of the **CPowerTime** object. |
| [Hour](#Hour) | Returns the Hour component of the **CPowerTime** object. It is a numeric value in the range of 0-23. |
| [IsFirstDayOfMonth](#IsFirstDayOfMonth) | Returns true if the date is the first day of the month; false, otherwise. |
| [IsLastDayOfMonth](#IsLastDayOfMonth) | Returns true if the date is the last day of the month; false, otherwise. |
| [IsLeapYear](#IsLeapYear) | Determines if a given year is a leap year or not. |
| [JulianToGregorian](#JulianToGregorian) | Converts a Julian date to a Gregorian date. |
| [Minute](#Minute) | Returns the Minute component of the **CPowerTime** object. This is a numeric value in the range of 0-59. |
| [Month](#Month) | Returns the Month component of the **CPowerTime** object. It is a  value in the range of 1-12. |
| [MonthString](#MonthString) | Returns the Month component of the **CPowerTime** object as a string. |
| [MSecond](#MSecond) | Returns the Millisecond component of the **CPowerTime** object.This is a numeric value in the range of 0-999. |
| [NewDate](#NewDate) | Sets a new date to this **CPowerTime** object. |
| [NewTime](#NewTime) | Sets a new time to this **CPowerTime** object. |
| [Now](#Now) | Assigns the current local date and time on this computer to this **CPowerTime** object. |
| [NowUTC](#NowUTC) | Assigns the current Coordinated Universal date and time (UTC) to this **CPowerTime** object. |
| [Second](#Second) | Returns the Second component of the **CPowerTime** object. This is a numeric value in the range of 0-59. |
| [SetFileTime](#SetFileTime) | Sets the date and time of this **CPowerTime** object. |
| [TimeString](#TimeString) | Retuns the time as a string based on the specified mask, e.g. "dd-MM-yyyy". |
| [Today](#Today) | Assigns the current local date on this computer to this **CPowerTime** object. |
| [UTCToLocal](#UTCToLocal) | Converts time based on the Coordinated Universal Time (UTC) to local file time. |
| [WeekOne](#WeekOne) | Returns the first day of the first week of the year. |
| [WeekNumber](#WeekNumber) | Returns the week number for a given date. |
| [WeeksInMonth](#WeeksInMonth) | Returns the number of weeks in the year. |
| [WeeksInYear](#WeeksInYear) | Returns the number of weeks in the specefied month and year. |
| [Year](#Year) | Returns the Year component of the **CPowerTime** object. |

# <a name="Constructors1"></a>Constructors

Create new **CPowerTime** objects initialized to the specified value.

```
CONSTRUCTOR CPowerTime
CONSTRUCTOR CPowerTime (BYVAL nTime AS ULONGLONG)
CONSTRUCTOR CPowerTime (BYREF ft AS FILETIME)
CONSTRUCTOR CPowerTime (BYREF st AS SYSTEMTIME)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTime* | A date and time expressed as a 64-bit value. |
| *ft* | A FILETIME structure. |
| *st* | A SYSTEMTIME structure. |

#### Examples

```
DIM cft AS CPowerTime = CPowerTime().GetCurrentTime()
```
```
DIM cft AS CPowerTime = AfxLocalFileTime
print cft.GetFileTime
```
```
DIM cft AS CPowerTime = AfxLocalSystemTime
print cft.GetFileTime
```

# <a name="CastOp1"></a>CAST Operator

Returns the **CPowerTime** value as an unsigned long integer.

```
OPERATOR CAST () AS ULONGLONG
```

#### Examples

```
DIM cft AS CPowerTime = CPowerTime().GetCurrentTime()
DIM nTime AS LONGLONG = cft
print nTime
```
```
DIM cft AS CPowerTime = CPowerTime().GetCurrentTime()
DIM cft2 AS CPowerTime = cft
```

# <a name="LetOp1"></a>LET Operator (=)

Assigns a value to a **CPowerTime** object.

```
OPERATOR LET (BYVAL nTime AS ULONGLONG)
OPERATOR LET (BYREF ft AS FILETIME)
OPERATOR LET (BYREF st AS SYSTEMTIME)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTime* | A date and time expressed as a 64-bit value. |
| *ft* | A FILETIME structure. |
| *st* | A SYSTEMTIME structure. |

#### Examples

```
DIM cft AS CPowerTime
cft = AfxLocalFileTime
print cft.GetFileTime
```
```
DIM cft AS CPowerTime
cft = AfxLocalSystemTime
print cft.GetFileTime
```

# <a name="Operators1"></a>Operators

Compares **CPowerTime** objects.

```
OPERATOR = (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR <> (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR < (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR > (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR <= (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
OPERATOR >= (BYREF dt1 AS CPowerTime, BYREF dt2 AS CPowerTime) AS BOOLEAN
```

# <a name="AddDays"></a>AddDays

Adds the specified number of days to this **CPowerTime** object. You can subtract days by using a negative number.

```
SUB AddDays (BYVAL nDays AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDays* | The number of days. |

# <a name="AddHours"></a>AddHours

Adds the specified number of hours to this **CPowerTime** object. The valid values for this member are 1 through 23. You can subtract hours by using a negative number.

```
AddHours (BYVAL nHours AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nHours* | The number of hours. |

# <a name="AddMinutes"></a>AddMinutes

Adds the specified number of minutes to this **CPowerTime** object. The valid values for this member are 1 through 59. You can subtract minutes by using a negative number.

```
SUB AddMinutes (BYVAL nMinutes AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMinutes* | The number of minutes. |

# <a name="AddMonths"></a>AddMonths

Adds the specified number of months to this **CPowerTime** object. The valid values for this member are 1 through 12. You can subtract months by using a negative number.

```
SUB AddMonths (BYVAL nMonths AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMonths* | The number of months. |

# <a name="AddMSeconds"></a>AddMSeconds

Adds the specified number of milliseconds to this **CPowerTime** object. The valid values for this member are 1 through 999. You can subtract milliseconds by using a negative number.

```
SUB AddSeconds (BYVAL nSeconds AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nSeconds* | The number of seconds. |

# <a name="AddSeconds"></a>AddSeconds

Adds the specified number of seconds to this **CPowerTime** object. The valid values for this member are 0 through 59. You can subtract seconds by using a negative number.

```
SUB AddSeconds (BYVAL nSeconds AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nSeconds* | The number of seconds. |

# <a name="AddYears"></a>AddYears

Adds the specified number of years to this **CPowerTime** object. You can subtract years by using a negative number.

```
SUB AddYears (BYVAL nYears AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYears* | The number of years. |

# <a name="AstroDay"></a>AstroDay

Returns the Astronomical Day for this **CPowerTime** object.

```
FUNCTION AstroDay () AS LONG
```

Returns the Astronomical Day for any given date.

```
FUNCTION AstroDay (BYVAL nYear AS LONG, BYVAL nMonth AS LONG, BYVAL nDay AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |
| *nMonth* | A month number (1-12). |
| *nDay* | A day number (1-31). |

#### Return value

The astronomical day.

#### Remarks

Among other things, can be used to find the number of days between any two dates, e.g.:

```
DIM cpt AS CPowerTime
PRINT cpt.AstroDay(-12400, 3, 1) - cpt.AstroDay(-12400, 2, 28)  ' Prints 2
PRINT cpt.AstroDay(12000, 3, 1) - cpt.AstroDay(-12000, 2, 28) ' Prints 8765822
PRINT cpt.AstroDay(1902, 2, 28) - cpt.AstroDay(1898, 3, 1)  ' Prints 1459 days
```

# <a name="AstroDayOfWeek"></a>AstroDayOfWeek

Calculates the day of the week, Sunday through Monday, of a given date.

```
FUNCTION AstroDayOfWeek (BYVAL nYear AS LONG = 0, BYVAL nMonth AS LONG = 0, BYVAL nDay AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |
| *nMonth* | A month number (1-12). |
| *nDay* | A day number (1-31). |

# <a name="Day"></a>Day

Returns the Day component of the **CPowerTime** object. It is a  value in the range of 1-31.

```
FUNCTION Day () AS LONG
```

# <a name="DaysDiff"></a>DaysDiff

Returns the days of difference between two dates.

The date part of the internal **CPowerTime** object is compared to the date part of the specified external **CPowerTime** object. The time-of-day part of each is ignored. The difference in number of days is returned as the result of the function.

```
FUNCTION DaysDiff (BYREF cpt AS CPowerTime) AS LONG
```
| Parameter  | Description |
| ---------- | ----------- |
| *cpt* | The **CPowerTime** object to compare. |

#### Example:

```
DIM cpt AS CPowerTime
cpt.NewDate(2019, 2, 2)
DIM cpt2 AS CPowerTime
cpt2.NewDate(1930, 12, 25)
print cpt.DaysDiff(cpt2)
' --or--
print cpt2.DaysDiff(cpt)
```

Calculates the days of difference between two dates.

```
FUNCTION DaysDiff (BYVAL nYear1 AS LONG, BYVAL nMonth1 AS LONG, BYVAL nDay1 AS LONG, _
   BYVAL nYear2 AS LONG, BYVAL nMonth2 AS LONG, BYVAL nDay2 AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear1* | A four digit year. |
| *nMonth1* | A month number (1-12). |
| *nDay1* | A day number (1-31). |
| *nYear2* | A four digit year. |
| *nMonth2* | A month number (1-12). |
| *nDay2* | A day number (1-31). |

#### Example:

```
DIM cpt AS CPowerTime
print cpt.DaysDiff(2019, 2, 2, 1930, 12, 25)
```

# <a name="DaysInMonth"></a>DaysInMonth

Returns the number of days of this **CPowerTime** object.

```
FUNCTION DaysInMonth () AS LONG
```

Returns the number of days in the specified month.

```
FUNCTION DaysInMonth (BYVAL nMonth AS LONG, BYVAL nYear AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMonth* | A month number (1-12). |
| *nYear* | A four digit year. |

# <a name="DayOfWeek"></a>DayOfWeek

Returns the Day-of-Week component of the **CPowerTime** object. It is a numeric value in the range of 0-6 (representing Sunday through Saturday).

```
FUNCTION DayOfWeek () AS LONG
```

# <a name="DayOfWeekString"></a>DayOfWeekString

Returns the Day-of-Week name of the **CPowerTime** object, expressed as a string (Monday, Tuesday...). The day name is appropriate for the locale, based upon the LCID parameter. If LCID is not given, the default LCID for the user is substituted.

```
FUNCTION DayOfWeekString (BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
```

# <a name="DayOfYear"></a>DayOfYear

Returns the day of the year, where Jan 1 is the first day of the year. If a parameter is omitted, the value stored in this **CPowerTime** object is assumed.

```
FUNCTION DayOfYear (BYVAL nYear AS LONG = 0, BYVAL nMonth AS LONG = 0, BYVAL nDay AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |
| *nMonth* | A month number (1-12). |
| *nDay* | A day number (1-31). |


# <a name="DaysInYear"></a>DaysInYear

Returns the number of days in the specified year. If the year is omitted, the year of this **CPowerTime** object is assumed.

```
FUNCTION DaysInYear (BYVAL nYear AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |

# <a name="Hour"></a>Hour

Returns the Hour component of the **CPowerTime** object. It is a numeric value in the range of 0-23.

```
FUNCTION Hour () AS LONG
```

# <a name="IsFirstDayOfMonth"></a>IsFirstDayOfMonth

Returns true if the date is the first day of the month; false, otherwise.

```
FUNCTION IsFirstDayOfMonth () AS BOOLEAN
```

# <a name="IsLastDayOfMonth"></a>IsLastDayOfMonth

Returns true if the date is the last day of the month; false, otherwise.

```
FUNCTION IsLastDayOfMonth () AS BOOLEAN
```

# <a name="IsLeapYear"></a>IsLeapYear

Determines if a given year is a leap year or not.

A leap year is defined as all years divisible by 4, except for years divisible by 100 that are not also divisible by 400. Years divisible by 400 are leap years. 2000 is a eap year. 1900 is not a leap year.

```
FUNCTION IsLeapYear (BYVAL nYear AS LONG = 0) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |

#### Return value

TRUE or FALSE.

# <a name="JulianToGregorian"></a>JulianToGregorian

Converts a Julian date to a Gregorian date.

```
FUNCTION JulianToGregorian (BYVAL nJulian AS LONG, BYVAL nDay AS LONG, _
   BYVAL nMonth AS LONG, BYVAL nYear AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nJulian* | The Julian date. |
| *nDay* | Out. The day (1-31). |
| *nMonth* | Out. The month (1-12). |
| *nYear* | Out. The four digit year. |

# <a name="Minute"></a>Minute

Returns the Minute component of the **CPowerTime** object. This is a numeric value in the range of 0-59.

```
FUNCTION Minute () AS LONG
```

# <a name="Month"></a>Month

Returns the Month component of the **CPowerTime** object. It is a  value in the range of 1-12.

```
FUNCTION Month () AS LONGç
```

# <a name="MonthString"></a>MonthString

Returns the Month component of the **CPowerTime** object, expressed as a string (January, February...).

```
FUNCTION MonthString () AS CWSTR
```

# <a name="MSecond"></a>MSecond

Returns the Millisecond component of the **CPowerTime** object. This is a numeric value in the range of 0-999.

```
FUNCTION MSecond () AS LONG
```

# <a name="Second"></a>Second

Returns the Second component of the **CPowerTime** object. This is a numeric value in the range of 0-59.

```
FUNCTION Second () AS LONG
```

# <a name="DateString"></a>DateString

Retuns the date as a string based on the specified mask, e.g. "dd-MM-yyyy".

```
FUNCTION DateString (BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMask* | A picture string that is used to form the date.<br>The format types "d", and "y" must be lowercase and the letter "M" must be uppercase.<br>For example, to get the date string "Wed, Aug 31 94", the application uses the picture string "ddd',' MMM dd yy". |
| *lcid* | Optional. The language identifier used for the conversion. Default is LOCALE_USER_DEFAULT. |

The following table defines the format types used to represent days.

| Format type | Meaning |
| ----------- | ----------- |
| d | Day of the month as digits without leading zeros for single-digit days. |
| dd | Day of the month as digits with leading zeros for single-digit days. |
| ddd | Abbreviated day of the week, for example, "Mon" in English (United States). |
| dddd | Day of the week. |

The following table defines the format types used to represent months.

| Format type | Meaning |
| ----------- | ----------- |
| M | Month as digits without leading zeros for single-digit months. |
| MM | Month as digits with leading zeros for single-digit months. |
| MMM | Abbreviated month, for example, "Nov" in English (United States). |
| MMMM | Month value, for example, "November" for English (United States), and "Noviembre" for Spanish (Spain). |

The following table defines the format types used to represent years.

| Format type | Meaning |
| ----------- | ----------- |
| y | Year represented only by the last digit. |
| yy | Year represented only by the last two digits. A leading zero is added for single-digit years. |
| yyyy | Year represented by a full four or five digits, depending on the calendar used. Thai Buddhist and Korean calendars have five-digit years. The "yyyy" pattern shows five digits for these two calendars, and four digits for all other supported calendars. Calendars that have single-digit or two-digit years, such as for the Japanese Emperor era, are represented differently. A single-digit year is represented with a leading zero, for example, "03". A two-digit year is represented with two digits, for example, "13". No additional leading zeros are displayed. |
| yyyyy | Behaves identically to "yyyy". |

#### Return value

The formatted date.


# <a name="TimeString"></a>TimeString

Retuns the time as a string based on the specified mask, e.g. "hh':'mm':'ss".

```
FUNCTION TimeString (BYREF wszMask AS WSTRING, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ft* | A FILETIME structure. |
| *wszMask* | A picture string that is used to form the time. |
| *lcid* | Optional. The language identifier used for the conversion. Default is LOCALE_USER_DEFAULT. |


The application can use the following elements to construct a format picture string. If spaces are used to separate the elements in the format string, these spaces appear in the same location in the output string. The letters must be in uppercase or lowercase as shown, for example, "ss", not "SS". Characters in the format string that are enclosed in single quotation marks appear in the same location and unchanged in the output string.

| Picture    | Meaning |
| ---------- | ----------- |
| h | Hours with no leading zero for single-digit hours; 12-hour clock |
| hh | Hours with leading zero for single-digit hours; 12-hour clock |
| H | Hours with no leading zero for single-digit hours; 24-hour clock |
| HH | Hours with leading zero for single-digit hours; 24-hour clock |
| m | Minutes with no leading zero for single-digit minutes |
| mm | Minutes with leading zero for single-digit minutes |
| s | Seconds with no leading zero for single-digit seconds |
| ss | Seconds with leading zero for single-digit seconds |
| t | One character time marker string, such as A or P |
| tt | Multi-character time marker string, such as AM or PM |

#### Return value

The formatted time.

# <a name="Format"></a>Format

Converts a **CPowerTime** object to a string.

```
FUNCTION Format (BYREF wszFmt AS WSTRING) AS CWSTR
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

#### Examples

```
print cft.Format("%A, %B %d, %Y %H:%M:%S")
```
```
DIM cft AS CPowerTime
cft = AfxLocalFileTime
print cft.Format("%A, %B %d, %Y %H:%M:%S")
```

# <a name="GetAsFileTime"></a>GetAsFileTime

Returns the date and time as a FILETIME structure.

```
FUNCTION GetAsFileTime () AS FILETIME
```

# <a name="GetAsJulianDate"></a>GetAsJulianDate

Returns the date as a Julian date.

```
FUNCTION GetAsJulianDate () AS LONG
```

Converts a Gregorian date to a Julian date. The year must be a 4 digit year.

```
FUNCTION GetAsJulianDate (BYVAL nYear AS LONG, BYVAL nMonth AS LONG, BYVAL nDay AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |
| *nMonth* | A month number (1-12). |
| *nDay* | A day number (1-31). |

# <a name="GetAsSystemTime"></a>GetAsSystemTime

Returns the date and time as a **SYSTEMTIME** structure.

```
FUNCTION GetAsSystemTime () AS SYSTEMTIME
```

# <a name="DateSerial"></a>DateSerial

Gets/sets the date and time as a FreeBasic date serial.

```
PROPERTY DateSerial () AS DOUBLE
PROPERTY DateSerial (BYVAL dTime AS DOUBLE)
```
#### Usage examples

```
DIM ct AS CPowerTime
ct.DateSerial = DateSerial(2019, 2, 4)
Print Format(ct.DateSerial, "yyyy/mm/dd") 
ct.DateSerial = DateValue("4/2/2019")
Print Format(ct.DateSerial, "yyyy/mm/dd") 
ct.DateSerial = TimeValue("11:59:59PM")
Print Format(dt, "hh:mm:ss")
ct.DateSerial = Now
Print Format(ct.DateSerial, "yyyy/mm/dd hh:mm:ss") 
ct.DateSerial = DateSerial(2005, 11, 28) + TimeSerial(7, 30, 50)
Print Format(ct.DateSerial, "yyyy/mm/dd hh:mm:ss") 
```

# <a name="GetCurrentTime"></a>GetCurrentTime

Returns a **CPowerTime** object that represents the current system date and time.

```
FUNCTION GetCurrentTime () AS CPowerTime
```

# <a name="GetFileTime"></a>GetFileTime

Returns the value of the **CPowerTime** object.

```
FUNCTION GetFileTime () AS LONGLONG
```

#### Example

```
DIM cft AS CPowerTime = CPowerTime().GetCurrentTime()
print cft.GetFileTime
```

# <a name="NewDate"></a>NewDate

Sets a new date to this **CPowerTime** object.

The date component of the **CPowerTime** object is assigned a new value based upon the specified parameters. The time component is unchanged. If parameters are invalid, **GetLastError** will return ERROR_INVALID_PARAMETER.

```
SUB NewDate (BYVAL nYear AS LONG = 0, BYVAL nMonth AS LONG = 0, BYVAL nDay AS LONG = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | The new year. The valid values for this member are 1601 through 30827. |
| *nMonth* | The new month. The valid values for this member are 1 through 12. |
| *nDay* | The new day. The valid values for this member are 1 through 31. |

# <a name="NewTime"></a>NewTime

Sets a new time to this **CPowerTime** object.

The time component of the **CPowerTime** object is assigned a new value based upon the specified parameters. The date component is unchanged. If parameters are invalid, **GetLastError** will return ERROR_INVALID_PARAMETER.

```
SUB NewTime (BYVAL nHour AS LONG = 0, BYVAL nMinute AS LONG = 0, _
    BYVAL nSecond AS LONG = 0, BYVAL nMSecond AS LONG = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nHour* | The new hour. The valid values for this member are 1 through 23. |
| *nMinute* | The new minute. The valid values for this member are 1 through 59. |
| *nSecond* | The new second. The valid values for this member are 1 through 59. |
| *nMSecond* | The new milliseond. The valid values for this member are 1 through 999. |

# <a name="Now"></a>Now

Assigns the current local date and time on this computer to this **CPowerTime** object.

```
SUB Now
```

# <a name="NowUTC"></a>NowUTC

Assigns the current Coordinated Universal date and time (UTC) to this **CPowerTime** object.

```
SUB NowUTC
```

# <a name="SetFileTime"></a>SetFileTime

Sets the date and time of this **CPowerTime** object

```
FUNCTION SetFileTime (BYVAL nTime AS ULONGLONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTime* | The date and time expressed as a 64-bit value. |

# <a name="Today"></a>Today

Assigns the current local date on this computer to this **CPowerTime** object.

```
SUB Today
```

# <a name="ToUTC"></a>ToUTC

The CPowerTime object is converted to Coordinated Universal Time (UTC). It is assumed that previous value was in local time.

```
SUB ToUTC
```

# <a name="ToLocalTime"></a>ToLocalTime

The CPowerTime object is converted to local time. It is assumed that the previous value was in Coordinated Universal Time (UTC).

```
SUB ToLocalTime
```

# <a name="WeekNumber"></a>WeekNumber

Returns the week number for a given date. The year must be a 4 digit year. If a parameter is omitted, the value stored in this **CPowerTime** object is assumed.

```
FUNCTION WeekNumber (BYVAL nYear AS LONG = 0, BYVAL nMonth AS LONG = 0, BYVAL nDay AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |
| *nMonth* | A month number (1-12). |
| *nDay* | A day number (1-31). |

# <a name="WeekOne"></a>WeekOne

Returns the first day of the first week of a year. If the year is omitted, the year of this **CPowerTime** object is assumed.

```
FUNCTION WeekOne (BYVAL nYear AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |

# <a name="WeeksInMonth"></a>WeeksInMonth

Returns the number of weeks in the specified month. Will be 4 or 5. If the year and month are omitted, the values of this **CPowerTime** object are assumed.

```
FUNCTION WeeksInMonth (BYVAL nMonth AS LONG = 0, BYVAL nYear AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nnMonth* | A month number (1-12). |
| *nYear* | A four digit year. |

# <a name="WeeksInYear"></a>WeeksInYear

Returns the number of weeks in the year, where weeks are taken to start on Monday. Will be 52 or 53. The year must be a four figit year. 
 If the year is omitted, the year of this **CPowertime** object is assumed.
```
FUNCTION WeeksInYear (BYVAL nYear AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nYear* | A four digit year. |

# <a name="Year"></a>Year

Returns the Year component of the **CPowerTime** object.

```
FUNCTION Year () AS LONG
```
