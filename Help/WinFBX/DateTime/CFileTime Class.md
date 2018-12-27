# CFileTime Class

This class provides methods for managing the date and time values associated with the creation, access and modification of files. The methods and data of this class are frequently used in conjunction with **CFileTimeSpan** objects, which deal with relative time values.

The date and time value is stored as a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601. This is the Coordinated Universal Time (UTC) format.

The following static const member variables are provided to simplify calculations (number of 100-nanosecond intervals):

```
CFileTime_Millisecond    10,000
CFileTime_Second         CFileTime_Millisecond * 1,000
CFileTime_Minute         CFileTime_Second * 60
CFileTime_Hour           CFileTime_Minute * 60
CFileTime_Day            CFileTime_Hour * 24
CFileTime_Week           CFileTime_Day * 7
```

**Note**: Not all file systems can record creation and last access time and not all file systems record them in the same manner. For example, on the Windows NT FAT file system, create time has a resolution of 10 milliseconds, write time has a resolution of 2 seconds, and access time has a resolution of 1 day (the access date). On NTFS, access time has a resolution of 1 hour. Furthermore, FAT records times on disk in local time, but NTFS records times on disk in UTC. 

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors1) | Create new **CFileTime** objects initialized to the specified value. |
| [CAST Operator](#CastOp1) | Returns the **CFileTime** value as a long integer. |
| [LET Operator](#LetOp1) | Assigns a value to a **CFileTime** object. |
| [Operators](#Operators1) | Adds, subtracts or compares **CFileTime** objects. |
| [Day](#Day) | Returns the number of 100-nanosecond intervals that make up one day. |
| [Format](#Format) | Converts a **CFileTime** object to a string. |
| [GetAsFileTime](#GetAsFileTime) | Returns the time as a **FILETIME** structure. |
| [GetAsSystemTime](#GetAsSystemTime) | Returns the time as a **SYSTEMTIME** structure. |
| [GetCurrentTime](#GetCurrentTime) | Returns a **CFileTime** object that represents the current system date and time. |
| [GetTime](#GetTime) | Returns the value of the **CFileTime** object. |
| [Hour](#Hour) | Returns the number of 100-nanosecond intervals that make up one hour. |
| [LocalToUTC](#LocalToUTC) | Converts a local file time to a file time based on the Coordinated Universal Time (UTC). |
| [Millisecond](#Millisecond) | Returns the number of 100-nanosecond intervals that make up one millisecond. |
| [Minute](#Minute) | Returns the number of 100-nanosecond intervals that make up one minute. |
| [Second](#Second) | Returns the number of 100-nanosecond intervals that make up one second. |
| [SetTime](#SetTime) | Sets the date and time of this **CFileTime** object. |
| [UTCToLocal](#UTCToLocal) | Converts time based on the Coordinated Universal Time (UTC) to local file time. |
| [Week](#Week) | Returns the number of 100-nanosecond intervals that make up one week. |

# CFileTimeSpan Class

This class provides methods for managing relative periods of time often encountered when performing operations concerning when a file was created, last accessed or last modified. The methods of this class are frequently used in conjunction with CFileTime class objects.

The CFileTimeSpan object can be created using an existing CFileTimeSpan object, or expressed as a 64-bit value. The default constructor sets the time span to 0.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#Constructors2) | Create new **CFileTimeSpan** objects initialized to the specified value. |
| [CAST Operator](#CastOp2) | Returns the **CFileTimeSpan** value as a long integer. |
| [LET Operator](#LetOp2) | Assigns a value to a **CFileTimeSpan** object. |
| [Operators](#Operators2) | Adds, subtracts or compares **CFileTimeSpan** objects. |
| [Day](#Day) | Returns the number of 100-nanosecond intervals that make up one day. |
| [GetTimeSpan](#GetTimeSpan) | Returns the value of the **CFileTimeSpan** object. |
| [Hour](#Hour) | Returns the number of 100-nanosecond intervals that make up one hour. |
| [Millisecond](#Millisecond) | Returns the number of 100-nanosecond intervals that make up one millisecond. |
| [Minute](#Minute) | Returns the number of 100-nanosecond intervals that make up one minute. |
| [Second](#Second) | Returns the number of 100-nanosecond intervals that make up one second. |
| [SetTimeSpan](#SetTimeSpan) | Sets the date and time of this **CFileTimeSpan** object. |
| [Week](#Week) | Returns the number of 100-nanosecond intervals that make up one week. |

# <a name="Constructors1"></a>Constructors (CFileTime)

Create new **CFileTime** objects initialized to the specified value.

```
CONSTRUCTOR CFileTime
CONSTRUCTOR CFileTime (BYVAL nSpan AS ULONGLONG)
CONSTRUCTOR CFileTime (BYREF ft AS FILETIME)
CONSTRUCTOR CFileTime (BYREF st AS SYSTEMTIME)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nSpan* | A date and time expressed as a 64-bit value. |
| *ft* | A FILETIME structure. |
| *st* | A SYSTEMTIME structure. |

#### Examples

```
DIM cft AS CFileTime = CFileTime().GetCurrentTime()
```
```
DIM cft AS CFileTime = AfxLocalFileTime
print cft.GetTime
```
```
DIM cft AS CFileTime = AfxLocalSystemTime
print cft.GetTime
```

# <a name="Constructors2"></a>Constructors (CFileTimeSpan)

Create new **CFileTimeSpan** objects initialized to the specified value.

```
CONSTRUCTOR CFileTimeSpan
CONSTRUCTOR CFileTimeSpan (BYVAL nSpan AS ULONGLONG)
CONSTRUCTOR CFileTimeSpan (BYREF cSpan AS CFileTimeSpan)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nSpan* | A period of time in milliseconds. |
| *cSpan* | A **CFileTimeSpan** object. |

#### Example

```
DIM cSpan AS CFileTimeSpan = CFileTime_Day
```

# <a name="CastOp1"></a>CAST Operator (CFileTime)

Returns the **CFileTime** value as a long integer.

```
OPERATOR CAST () AS LONGLONG
```

#### Examples

```
DIM cft AS CFileTime = CFileTime().GetCurrentTime()
DIM nTime AS LONGLONG = cft
print nTime
```
```
DIM cft AS CFileTime = CFileTime().GetCurrentTime()
DIM cft2 AS CFileTime = cft
```

# <a name="CastOp2"></a>CAST Operator (CFileTimeSpan)

Returns the **CFileTimeSpan** value as a long integer.

```
OPERATOR CAST () AS LONGLONG
```

# <a name="LetOp1"></a>LET Operator (=) (CFileTime)

Assigns a value to a **CFileTime** object.

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
DIM cft AS CFileTime
cft = AfxLocalFileTime
print cft.GetTime
```
```
DIM cft AS CFileTime
cft = AfxLocalSystemTime
print cft.GetTime
```

# <a name="LetOp2"></a>LET Operator (=) (CFileTimeSpan)

Assigns a value to a **CFileTimeSpan** object.

```
OPERATOR LET (BYVAL nSpan AS LONGLONG)
OPERATOR LET (BYREF cSpan AS CFileTimeSpan)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nSpan* | A period of time in milliseconds. |
| *cSpan* | A **CFileTimeSpan** object. |

# <a name="Operators1"></a>Operators (CFileTime)

Adds, subtracts or compares **CFileTime** objects.

```
OPERATOR + (BYREF dt AS CFileTime, BYREF dateSpan AS CFileTimeSpan) AS CFileTime
OPERATOR - (BYREF dt AS CFileTime, BYREF dateSpan AS CFileTimeSpan) AS CFileTime
OPERATOR - (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS CFileTimeSpan
OPERATOR += (BYREF dateSpan AS CFileTimeSpan)
OPERATOR -= (BYREF dateSpan AS CFileTimeSpan)
OPERATOR = (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
OPERATOR <> (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
OPERATOR < (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
OPERATOR > (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
OPERATOR <= (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
OPERATOR >= (BYREF dt1 AS CFileTime, BYREF dt2 AS CFileTime) AS BOOLEAN
```

# <a name="Operators2"></a>Operators (CFileTimeSpan)

Adds, subtracts or compares **CFileTimeSpan** objects.

```
OPERATOR + (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS CFileTimeSpan
OPERATOR - (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS CFileTimeSpan
OPERATOR += (BYREF dateSpan AS CFileTimeSpan)
OPERATOR -= (BYREF dateSpan AS CFileTimeSpan)
OPERATOR = (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
OPERATOR <> (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
OPERATOR < (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
OPERATOR > (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
OPERATOR <= (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
OPERATOR >= (BYREF cSpan1 AS CFileTimeSpan, BYREF cSpan2 AS CFileTimeSpan) AS BOOLEAN
```

# <a name="Day"></a>Day (CFileTime / CFileTimeSpan)

Returns the number of 100-nanosecond intervals that make up one day.

```
FUNCTION Day () AS ULONGLONG
```

# <a name="Hour"></a>Hour (CFileTime / CFileTimeSpan)

Returns the number of 100-nanosecond intervals that make up one hour.

```
FUNCTION Hour () AS ULONGLONG
```

# <a name="Millisecond"></a>Millisecond (CFileTime / CFileTimeSpan)

Returns the number of 100-nanosecond intervals that make up one millisecond.

```
FUNCTION Millisecond () AS ULONGLONG
```

# <a name="Minute"></a>Minute (CFileTime / CFileTimeSpan)

Returns the number of 100-nanosecond intervals that make up one minute.

```
FUNCTION Minute () AS ULONGLONG
```

# <a name="Second"></a>Second (CFileTime / CFileTimeSpan)

Returns the number of 100-nanosecond intervals that make up one second.

```
FUNCTION Second () AS ULONGLONG
```

# <a name="Week"></a>Week (CFileTime / CFileTimeSpan)

Returns the number of 100-nanosecond intervals that make up one week.

```
FUNCTION Week () AS ULONGLONG
```

# <a name="Format"></a>Format (CFileTime)

Converts a **CFileTime** object to a string.

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
DIM cft AS CFileTime
cft = AfxLocalFileTime
print cft.Format("%A, %B %d, %Y %H:%M:%S")
```

# <a name="GetAsFileTime"></a>GetAsFileTime (CFileTime)

Returns the time as a FILETIME structure.

```
FUNCTION GetAsFileTime () AS FILETIME
```

# <a name="GetAsSystemTime"></a>GetAsSystemTime (CFileTime)

Returns the time as a **SYSTEMTIME** structure.

```
FUNCTION GetAsSystemTime () AS SYSTEMTIME
```

# <a name="GetCurrentTime"></a>GetCurrentTime (CFileTime)

Returns a **CFileTime** object that represents the current system date and time.

```
FUNCTION GetCurrentTime () AS CFileTime
```

# <a name="GetTime"></a>GetTime (CFileTime)

Returns the value of the **CFileTime** object.

```
FUNCTION GetTime () AS LONGLONG
```

#### Example

```
DIM cft AS CFileTime = CFileTime().GetCurrentTime()
print cft.GetTime
```

# <a name="SetTime"></a>SetTime (CFileTime)

Sets the date and time of this **CFileTime** object

```
FUNCTION SetTime (BYVAL nTime AS ULONGLONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTime* | The date and time expressed as a 64-bit value. |

# <a name="LocalToUTC"></a>LocalToUTC (CFileTime)

Converts a local file time to a file time based on the Coordinated Universal Time (UTC).

```
FUNCTION LocalToUTC () AS CFileTime
```

# <a name="UTCToLocal"></a>UTCToLocal (CFileTime)

Converts time based on the Coordinated Universal Time (UTC) to local file time.

```
FUNCTION UTCToLocal () AS CFileTime
```

# <a name="GetTimeSpan"></a>GetTimeSpan (CFileTimeSpan)

Returns the value of the **CFileTimeSpan** object.

```
FUNCTION GetTimeSpan () AS LONGLONG
```

# <a name="SetTimeSpan"></a>SetTimeSpan (CFileTimeSpan)

Sets the date and time of this **CFileTimeSpan** object.

```
FUNCTION SetTimeSpan (BYVAL nSpan AS LONGLONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nSpan* | The new value for the time span in milliseconds. |
