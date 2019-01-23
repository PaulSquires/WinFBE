# String Procedures

**Include file**: AfxStr.inc

| Name       | Description |
| ---------- | ----------- |
| [AfxStrClipLeft](#AfxStrClipLeft) | Returns a string with the specified number of characters removed from the left side of the string. |
| [AfxStrClipMid](#AfxStrClipMid) | Returns a string with the specified number of characters removed starting at the specified position. |
| [AfxStrClipRight](#AfxStrClipRight) | Returns a string with the specified number of characters removed from the right side of the string. |
| [AfxStrCSet](#AfxStrCSet) | Returns a string containing a centered string. |
| [AfxStrDelete](#AfxStrDelete) | Deletes a specified number of characters from a string expression. |
| [AfxStrExtract](#AfxStrExtract) | Extracts characters from a string up to (but not including) the specified matching. Case sensitive. |
| [AfxStrExtractI](#AfxStrExtractI) | Extracts characters from a string up to (but not including) the specified matching string. Case insensitive. |
| [AfxStrExtractAny](#AfxStrExtractAny) | Extracts characters from a string up to (but not including) any character in the matching string. Case sensitive. |
| [AfxStrExtractAnyI](#AfxStrExtractAnyI) | Extracts characters from a string up to (but not including) any character in the matching string. Case insensitive. |
| [AfxStrInsert](#AfxStrInsert) | Inserts a string at a specified position within another string expression. |
| [AfxStrJoin](#AfxStrJoin) | Returns a string consisting of all of the strings in an array, each separated by a delimiter. |
| [AfxStrLCase](#AfxStrLCase) | Returns a lowercased version of a string. |
| [AfxStrLSet](#AfxStrLSet) | Returns a string containing a left justified string. |
| [AfxStrParse](#AfxStrParse) | Returns a delimited field from a string expression. |
| [AfxStrParseAny](#AfxStrParseAny) | Returns a delimited field from a string expression. Supports more than one character for the delimiter. |
| [AfxStrParseCount](#AfxStrParseCount) | Returns the count of delimited fields from a string expression. |
| [AfxStrParseCountAny](#AfxStrParseCountAny) | Returns the count of delimited fields from a string expression. Supports more than one character for the delimiter. |
| [AfxStrRemain](#AfxStrRemain) | Returns the portion of a string following the first occurrence of a string. Case sensitive. |
| [AfxStrRemainI](#AfxStrRemainI) | Returns the portion of a string following the first occurrence of a string. Case insensitive. |
| [AfxStrRemainAny](#AfxStrRemainAny) | Returns the portion of a string following the first occurrence of a group of characters. Case sensitive. |
| [AfxStrRemainAnyI](#AfxStrRemainAnyI) | Returns the portion of a string following the first occurrence of a group of characters. Case insensitive. |
| [AfxStrRemove](#AfxStrRemove) | Returns a new string with substrings removed. Case sesnsitive. |
| [AfxStrRemoveI](#AfxStrRemoveI) | Returns a new string with substrings removed. Case insensitive. |
| [AfxStrRemoveAny](#AfxStrRemoveAny) | Returns a new string with characters removed. Case sesnsitive. |
| [AfxStrRemoveAnyI](#AfxStrRemoveAnyI) | Returns a new string with characters removed. Case insesnsitive. |
| [AfxStrRepeat](#AfxStrRepeat) | Returns a string consisting of multiple copies of the specified string. |
| [AfxStrReplace](#AfxStrReplace) | Replaces all the occurrences of a string with another string. Case sensitive. |
| [AfxStrReplaceI](#AfxStrReplaceI) | Replaces all the occurrences of a string with another string. Case insensitive. |
| [AfxStrReplaceAny](#AfxStrReplaceAny) | Replaces all the occurrences of a group of characters with another character. Case sensitive. |
| [AfxStrReplaceAnyI](#AfxStrReplaceAnyI) | Replaces all the occurrences of a group of characters with another character. Case insensitive. |
| [AfxStrRetain](#AfxStrRetain) | Returns a string containing only the characters contained in a specified match string. Case sensitive. |
| [AfxStrRetainI](#AfxStrRetainI) | Returns a string containing only the characters contained in a specified match string. Case insensitive. |
| [AfxStrRetainAny](#AfxStrRetainAny) | Returns a string containing only the characters contained in a specified group of characters. Case sensitive. |
| [AfxStrRetainAnyI](#AfxStrRetainAnyI) | Returns a string containing only the characters contained in a specified group of characters. Case insensitive. |
| [AfxStrReverse](#AfxStrReverse) | Reverses the contents of a string expression. |
| [AfxStrRSet](#AfxStrRSet) | Returns a string containing a right justified string. |
| [AfxStrShrink](#AfxStrShrink) | Shrinks a string to use a consistent single character delimiter. |
| [AfxStrSplit](#AfxStrSplit) | Splits a string into tokens, which are sequences of contiguous characters separated by any of the characters that are part of delimiters. |
| [AfxStrSpn](#AfxStrSpn) | Returns the length of the initial portion of a string which consists only of characters that are part of a specified set of characters. |
| [AfxStrTally](#AfxStrTally) | Count the number of occurrences of a string within a string. Case sensitive. |
| [AfxStrTallyI](#AfxStrTallyI) | Count the number of occurrences of a string within a string. Case insensitive. |
| [AfxStrTallyAny](#AfxStrTallyAny) | Count the number of occurrences of a list of characters within a string. Case sensitive. |
| [AfxStrTallyAnyI](#AfxStrTallyAnyI) | Count the number of occurrences of a list of characters within a string. Case insensitive. |
| [AfxStrUCase](#AfxStrUCase) | Returns an uppercased version of a string. |
| [AfxStrVerify](#AfxStrVerify) | Determine whether each character of a string is present in another string. Case sensitive. |
| [AfxStrVerifyI](#AfxStrVerifyI) | Determine whether each character of a string is present in another string. Case insensitive. |
| [AfxStrWrap](#AfxStrWrap) | Adds paired characters to the beginning and end of a string. |
| [AfxStrUnWrap](#AfxStrUnWrap) | Removes paired characters to the beginning and end of a string. |
| [AfxStrPathName](#AfxStrPathName) | Parses a path to extract component parts. |
| [AfxStrFormatByteSize](#AfxStrFormatByteSize) | Converts a numeric value into a string that represents the number expressed as a size value in bytes, kilobytes, megabytes, or gigabytes, depending on the size. |
| [AfxStrFormatKBSize](#AfxStrFormatKBSize) | Converts a numeric value into a string that represents the number expressed as a size value in kilobytes. |
| [AfxStrFromTimeInterval](#AfxStrFromTimeInterval) | Converts a time interval, specified in milliseconds, to a string. |
| [AfxBase64DecodeA](#AfxBase64DecodeA) | Converts the contents of a Base64 mime encoded string to an ascii string. |
| [AfxBase64DecodeW](#AfxBase64DecodeW) | Converts the contents of a Base64 mime encoded string to an unicode string. |
| [AfxBase64EncodeA](#AfxBase64EncodeA) | Converts the contents of an ascii string to Base64 mime encoding. |
| [AfxBase64EncodeW](#AfxBase64EncodeW) | Converts the contents of an unicode string to Base64 mime encoding. |
| [AfxCryptBinaryToString](#AfxCryptBinaryToString) | Converts an array of bytes into a formatted string. |
| [AfxCryptStringToBinary](#AfxCryptStringToBinary) | Converts a formatted string into an array of bytes. |
# <a name="AfxStrClipLeft"></a>AfxStrClipLeft

Returns a string with *nCount* characters removed from the left side of the string.

```
FUNCTION AfxStrClipLeft (BYREF wszMainStr AS CONST WSTRING, BYVAL nCount AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *nCount* | The number of characters to be removed. |

# <a name="AfxStrClipMid"></a>AfxStrClipMid

Returns a string with *nCount* characters removed from the left side of the string.

```
FUNCTION AfxStrClipMid (BYREF wszMainStr AS CONST WSTRING, BYVAL nStart AS LONG, BYVAL nCount AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *nStart* | The one-based starting position. |
| *nCount* | The number of characters to be removed. |

# <a name="AfxStrClipRight"></a>AfxStrClipRight

Returns a string with *nCount* characters removed from the right side of the string.

```
FUNCTION AfxStrClipRight  (BYREF wszMainStr AS CONST WSTRING, BYVAL nCount AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *nCount* | The number of characters to be removed. |

# <a name="AfxStrCSet"></a>AfxStrCSet

Returns a string containing a centered string.

```
FUNCTION AfxStrCSet (BYREF wszMainStr AS CONST WSTRING, _
   BYVAL nStringLength AS LONG, BYREF wszPadCharacter AS CONST WSTRING = " ") AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The string to be justified. |
| *nStringLength* | The length of the new string. |
| *wszPadCharacter* | The character to be used for padding. If it is not specified, the string will be padded with spaces. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrCSet("FreeBasic", 20, "*")
```

# <a name="AfxStrDelete"></a>AfxStrDelete

Deletes a specified number of characters from a string expression.

```
FUNCTION AfxStrDelete (BYREF wszMainStr AS CONST WSTRING, _
   BYVAL nStart AS LONG, BYVAL nCount AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *nStart* | The one-based starting position. |
| *nCount* | The number of characters to be removed. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrDelete("1234567890", 4, 3)   ' Returns 1234890"
```

# <a name="AfxStrExtract"></a>AfxStrExtract

Extracts characters from a string up to (but not including) a string. Case sensitive.

```
FUNCTION AfxStrExtract (BYVAL nStart AS LONG, _
   BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The one-based starting position. |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string to be searched. |

#### Usage example

```
The following line returns "aba" (match on "cad")
DIM cws AS CWSTR = AfxStrExtract(1, "abacadabra","cad")
```

# AfxStrExtract (Overload)

```
FUNCTION AfxStrExtract (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszDelim1 AS WSTRING, BYREF wszDelim2 AS WSTRING) AS CWSTR
```

Extracts the portion of a string following the occurrence of a specified delimiter up to the second delimiter. If one or both of the delimiters aren't found, it returns an empty string.

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszDelim1* | The first delimiter. |
| *wszDelim2* | The second delimiter. |

#### Usage example

```
The following lines return "text between parentheses" (text delimited by "(" and ")")
DIM cws AS CWSTR = "blah blah (text beween parentheses) blah blah"
PRINT AfxStrExtract(cws, "(", ")")
```

# AfxStrExtract (Overload)

```
FUNCTION AfxStrExtract (BYVAL nStart AS LONG, BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszDelim1 AS WSTRING, BYREF wszDelim2 AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The one-based starting position. |
| *wszMainStr* | The main string. |
| *wszDelim1* | The first delimiter. |
| *wszDelim2* | The second delimiter. |

# <a name="AfxStrExtractI"></a>AfxStrExtractI

Extracts characters from a string up to (but not including) a string. Case insensitive.

```
FUNCTION AfxStrExtractI (BYVAL nStart AS LONG, _
   BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The one-based starting position. |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string to be searched. |

#### Usage example

```
The following line returns "aba" (match on "CaD")
DIM cws AS CWSTR = AfxStrExtractI(1, "abacadabra","CaD")
```
# <a name="AfxStrExtractAny"></a>AfxStrExtractAny

Extracts characters from a string up to (but not including) a group of characters. Case sensitive.

```
FUNCTION AfxStrExtractAny (BYVAL nStart AS LONG, _
   BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The one-based starting position. |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The characters to be searched individually. A match on any one of which will cause the extract operation to be performed up to that character. |

#### Usage example

```
The following line returns "aba" (match on "c")
DIM cws AS CWSTR = AfxStrExtractAny(1, "abacadabra","cd")
```
# <a name="AfxStrExtractAnyI"></a>AfxStrExtractAnyI

Extracts characters from a string up to (but not including) a group of characters. Case insensitive.

```
FUNCTION AfxStrExtractAnyI (BYVAL nStart AS LONG, _
   BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The one-based starting position. |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The characters to be searched individually. A match on any one of which will cause the extract operation to be performed up to that character. |

#### Usage example

```
The following line returns "aba" (match on "c")
DIM cws AS CWSTR = AfxStrExtractAnyI(1, "abacadabra","Cd")
```

# <a name="AfxStrFormatByteSize"></a>AfxStrFormatByteSize

Converts a numeric value into a string that represents the number expressed as a size value in bytes, kilobytes, megabytes, or gigabytes, depending on the size.

```
FUNCTION AfxStrFormatByteSize (BYVAL ull AS ULONGLONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ull* | The numeric value to be converted. |

# <a name="AfxStrFormatKBSize"></a>AfxStrFormatKBSize

Converts a numeric value into a string that represents the number expressed as a size value in kilobytes.

```
FUNCTION AfxStrFormatKBSize (BYVAL ull AS ULONGLONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *ull* | The numeric value to be converted. |

# <a name="AfxStrFromTimeInterval"></a>AfxStrFromTimeInterval

Converts a time interval, specified in milliseconds, to a string.

```
FUNCTION AfxStrFromTimeInterval (BYVAL dwTimeMS AS DWORD, BYVAL digits AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwTimeMS* | The time interval, in milliseconds. |
| *digits* | The maximum number of significant digits to be represented in the output string. |

Some examples for *digits*:

| dwTimeMS   | digits      | cwsOut      |
| ---------- | ----------- | ----------- |
| 34000 | 3 | 34 sec |
| 34000 | 2 | 34 sec |
| 34000 | 1 | 30 sec |
| 74000 | 3 | 1 min 14 sec |
| 74000 | 2 | 1 min 10 sec |
| 74000 | 2 | 1 min |

# <a name="AfxStrInsert"></a>AfxStrInsert

Inserts a string at a specified position within another string expression.

```
FUNCTION AfxStrInsert (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszInsertString AS CONST WSTRING, BYVAL nPosition AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszInsertString* | The string to be inserted. |
| *nPosition* | The one-based starting position. If *nPosition* is greater than the length of *wszMainStr* or <= zero then *wszInsertString* is appended to *wszMainStr*. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrInsert("1234567890", "--", 6)   ' Returns "123456--7890"
```

# <a name="AfxStrJoin"></a>AfxStrJoin

Returns a string consisting of all of the strings in an array, each separated by a delimiter. If the delimiter is a null (zero-length) string then no separators are inserted between the string sections. If the delimiter expression is the 3-byte value of "," which may be expressed in your source code as the string literal """,""" or as Chr(34,44,34) then a leading and trailing double-quote is added to each string section. This ensures that the returned string contains standard comma-delimited quoted fields that can be easily parsed.

```
FUNCTION AfxStrJoin (BYREF cwsa AS CSafeArray, BYREF wszDelimiter AS WSTRING = " ") AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsa* | The one-dimensional VT_BSTR CSafeArray to join. |
| *wszDelimiter* | The delimiter character. |

#### Return value

A CWSTR containing the joined string.

#### Usage example

```
DIM csa AS CSafeArray = CSafeArray("STRING", 3, 1)
csa.PutStr(1, "One")
csa.PutStr(2, "Two")
csa.PutStr(3, "Three")
DIM cws AS CWSTR = AfxStrJoin(csa, ",")
PRINT cws   ' ouput: One,Two,Three
```
**Include file**: CSafeArray.inc

# <a name="AfxStrLCase"></a>AfxStrLCase

Returns a lowercased version of a string.

```
FUNCTION AfxStrLCase (BYVAL pwszStr AS WSTRING PTR, _
   BYVAL pwszLocaleName AS WSTRING PTR = LOCALE_NAME_USER_DEFAULT, _
   BYVAL dwMapFlags AS DWORD = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszStr* | The string. |
| *pwszLocaleName* | Optional. Pointer to a locale name or one of these pre-defined values: LOCALE_NAME_INVARIANT, LOCALE_NAME_SYSTEM_DEFAULT, LOCALE_NAME_USER_DEFAULT |
| *dwMapFlags* | Optional. Flag specifying the type of transformation to use during string mapping or the type of sort key to generate. |

For a table of language culture names see: [Table of Language Culture Names, Codes, and ISO Values](https://docs.microsoft.com/en-us/previous-versions/commerce-server/ee825488(v=cs.20))

For a complete list see: [LCMapStringEx function](https://docs.microsoft.com/en-us/windows/desktop/api/winnls/nf-winnls-lcmapstringex)

#### Remarks

The string conversion functions available in FreeBasic are not fully suitable for some languages. For example, the Turkish word "karışıklığı" is uppercased as "KARıŞıKLıĞı" instead of "KARIŞIKLIĞI", and "KARIŞIKLIĞI" is lowercased to "karişikliği" instead of "karışıklığı". Notice the "ı", that is not an "i".

#### Return value

The lowercased string.

# <a name="AfxStrLSet"></a>AfxStrLSet

Returns a string containing a left justified string.

```
FUNCTION AfxStrLSet (BYREF wszMainStr AS CONST WSTRING, _
   BYVAL nStringLength AS LONG, BYREF wszPadCharacter AS CONST WSTRING = " ") AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The string to be justified. |
| *nStringLength* | The length of the new string. |
| *wszPadCharacter* | The character to be used for padding. If it is not specified, the string will be padded with spaces. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrLSet("FreeBasic", 20, "*")
```

# <a name="AfxStrParse"></a>AfxStrParse 

Returns a delimited field from a string expression.

```
FUNCTION AfxStrParse (BYREF wszMainStr AS CONST WSTRING, _
   BYVAL nPosition AS LONG = 1, BYREF wszDelimiter AS CONST WSTRING = ",") AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The string to be parsed. |
| *nPosition* | Starting position. If *nPosition* is zero or is outside of the actual field count, an empty string is returned. If *nPosition* is negative, fields are searched from the right to left of the *wszMainStr*. |
| *wszDelimiter* | A string of one or more characters that must be fully matched to be successful. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrParse("one,two,three", 2)   ' Returns "two"
DIM cws AS CWSTR = AfxStrParse("one;two,three", 1, ";")   ' Returns "one"
```

# <a name="AfxStrParseAny"></a>AfxStrParseAny 

Returns a delimited field from a string expression.

```
FUNCTION AfxStrParse (BYREF wszMainStr AS CONST WSTRING, _
   BYVAL nPosition AS LONG = 1, BYREF wszDelimiter AS CONST WSTRING = ",") AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The string to be parsed. |
| *nPosition* | Starting position. If *nPosition* is zero or is outside of the actual field count, an empty string is returned. If *nPosition* is negative, fields are searched from the right to left of the *wszMainStr*. |
| *wszDelimiter* | A string of one or more characters any of which may act as a delimiter character. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrParseAny("1;2,3", 2, ",;")   ' Returns "2"
```

# <a name="AfxStrParseCount"></a>AfxStrParseCount 

Returns the count of delimited fields from a string expression.

```
FUNCTION AfxStrParseCount (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszDelimiter AS CONST WSTRING = ",") AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. If *wszMainStr* is empty (a null string) or contains no delimiter character(s), the string is considered to contain exactly one subfield. |
| *wszDelimiter* | One or more character delimiters that must be fully matched. Delimiters are case-sensitive. |

#### Usage example

```
DIM nCount AS LONG = AfxStrParseCount("one,two,three", ",")   ' Returns 3
```

# <a name="AfxStrParseCountAny"></a>AfxStrParseCountAny 

Returns the count of delimited fields from a string expression.

```
FUNCTION AfxStrParseCountAny (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszDelimiter AS CONST WSTRING = ",") AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. If wszMainStr is empty (a null string) or contains no delimiter character(s), the string is considered to contain exactly one sub-field. |
| *wszDelimiter* | A set of characters (one or more), any of which may act as a delimiter character. Delimiters are case-sensitive. |

#### Usage example

```
DIM nCount AS LONG = AfxStrParseCountAny("1;2,3", ",;")   ' Returns 3
```

# <a name="AfxStrPathName"></a>AfxStrPathName 

Parses a path to extract component parts.

```
FUNCTION AfxStrPathName (BYREF wszOption AS CONST WSTRING, BYREF wszFileSpec AS WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszOption* | One of the following words which is used to specify the requested part: PATH, NAME, EXTN, NAMEX. |
| *wszFileSpec* | The path to be scanned. |

| Keyword    | Action      |
| ---------- | ----------- |
| **PATH** | Returns the path portion of the path/file Name. That is the text up to and including the last backslash (\) or colon (:). |
| **NAME** | Returns the name portion of the path/file Name. That is the text to the right of the last backslash (\) or colon (:), ending just before the last period (.). |
| **EXTN** | Returns the extension portion of the path/file name. That is the last period (.) in the string plus the text to the right of it. |
| **NAMEX** | Returns the name and the extension parts combined. |

# <a name="AfxStrRemain"></a>AfxStrRemain

Returns the portion of a string following the first occurrence of a string. Case sensitive.

```
FUNCTION AfxStrRemain (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string to search for. |
| *nStart* | Optional. Starting position to begin the search. If *nStart* is not specified, the search will begin at position 1. If nStart is zero, a nul string is returned. If *nStart* is negative, the starting position is counted from right to left: -1 for the last character, -2 for the second to last, etc. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRemain("Brevity is the soul of wit", "is ")   ' Returns "the soul of wit"
```

# <a name="AfxStrRemainI"></a>AfxStrRemainI

Returns the portion of a string following the first occurrence of a string. Case insensitive.

```
FUNCTION AfxStrRemainI (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string to search for. |
| *nStart* | Optional. starting position to begin the search. If *nStart* is not specified, the search will begin at position 1. If nStart is zero, a nul string is returned. If *nStart* is negative, the starting position is counted from right to left: -1 for the last character, -2 for the second to last, etc. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRemain("Brevity is the soul of wit", "Is ")   ' Returns "the soul of wit"
```

# <a name="AfxStrRemainAny"></a>AfxStrRemainAny

Returns the portion of a string following the first occurrence of a list of characters. Case sensitive.

```
FUNCTION AfxStrRemainAny (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The characters to search for. |
| *nStart* | Optional. starting position to begin the search. If *nStart* is not specified, the search will begin at position 1. If nStart is zero, a nul string is returned. If *nStart* is negative, the starting position is counted from right to left: -1 for the last character, -2 for the second to last, etc. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRemainAny("I think, therefore I am", ",")   ' Returns "therefore I am"
```

# <a name="AfxStrRemainAnyI"></a>AfxStrRemainAnyI

Returns the portion of a string following the first occurrence of a list of characters. Case insensitive.

```
FUNCTION AfxStrRemainAnyI (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING, BYVAL nStart AS LONG = 1) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The characters to search for. |
| *nStart* | Optional. starting position to begin the search. If *nStart* is not specified, the search will begin at position 1. If nStart is zero, a nul string is returned. If *nStart* is negative, the starting position is counted from right to left: -1 for the last character, -2 for the second to last, etc. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRemainAnyI("I think, therefore I am", "E")   ' -> "refore I am"
```

# <a name="AfxStrRemove"></a>AfxStrRemove

Returns a new string with strings removed. Case sensitive.

```
FUNCTION AfxStrRemove (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string expression to be removed. If *wszMatchStr* is not present in *wszMainStr*, all of *wszMainStr* is returned intact. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRemove("Hello World. Welcome to the Freebasic World", "World")
```

# AfxStrRemove (Overload)

Returns a copy of a string with a substring enclosed between the specified delimiters removed. Case sensitive.

```
FUNCTION AfxStrRemove (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszDelim1 AS CONST WSTRING, BYREF wszDelim2 AS CONST WSTRING, _
   BYVAL fRemoveAll AS BOOLEAN = FALSE) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszDelim1* | The first delimiter |
| *wszDelim2* | The second delimiter |
| *fRemoveAll* | Recursively remove all the occurrences. |

#### Usage examples

```
DIM cwsText AS CWSTR = "blah blah (text beween parentheses) blah blah"
DIM cws AS CWSTR = AfxStrRemove(cwsText, "(", ")")   ' Returns "blah blah  blah blah"
```

```
DIM cwsText AS CWSTR = "As Long var1(34), var2(  73 ), var3(any)"
DIM cws AS CWSTR = AfxStrRemove(cwsText, "(", ")", TRUE)   ' Returns "As Long var1, var2, var3"
```

# AfxStrRemove (Overload)

Returns a copy of a string with a substring enclosed between the specified delimiters removed. Case sensitive.

```
FUNCTION AfxStrRemove (BYVAL nSart AS LONG = 1, BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszDelim1 AS CONST WSTRING, BYREF wszDelim2 AS CONST WSTRING, _
   BYVAL fRemoveAll AS BOOLEAN = FALSE) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nSart* | Optional. The one-based starting position where to start the search. |
| *wszMainStr* | The main string. |
| *wszDelim1* | The first delimiter |
| *wszDelim2* | The second delimiter |
| *fRemoveAll* | Recursively remove all the occurrences. |

#### Usage examples

```
DIM cwsText AS CWSTR = "blah blah (text beween parentheses) blah blah"
DIM cws AS CWSTR = AfxStrRemove(cwsText, "(", ")")   ' Returns "blah blah  blah blah"
```

# <a name="AfxStrRemoveI"></a>AfxStrRemoveI

Returns a new string with strings removed. Case insensitive.

```
FUNCTION AfxStrRemoveI (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string expression to be removed. If *wszMatchStr* is not present in *wszMainStr*, all of *wszMainStr* is returned intact. |

#### Usage example

```
AfxStrRemoveI("Hello World. Welcome to the Freebasic World", "world")
```

# <a name="AfxStrRemoveAny"></a>AfxStrRemoveAny

Returns a new string with characters removed. Case sensitive.

```
FUNCTION AfxStrRemoveAny (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string expression to be removed. If *wszMatchStr* is not present in *wszMainStr*, all of *wszMainStr* is returned intact. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRemoveAny("abacadabra", "bac")   ' -> "dr"
```

# <a name="AfxStrRemoveAnyI"></a>AfxStrRemoveAnyI

Returns a new string with characters removed. Case insensitive.

```
FUNCTION AfxStrRemoveAnyI (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string expression to be removed. If *wszMatchStr* is not present in *wszMainStr*, all of *wszMainStr* is returned intact. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRemoveAnyI("abacadabra", "BaC")   ' -> "dr"
```

# <a name="AfxStrRepeat"></a>AfxStrRepeat

Returns a string consisting of multiple copies of the specified string. This function is similar to STRING, but STRING only makes multiple copies of a single character.

```
FUNCTION AfxStrRepeat (BYVAL nCount AS LONG, BYREF wszStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nCount* | The number of copies. |
| *wszStr* | The string to be copied. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRepeat(5, "Paul")
```

# <a name="AfxStrReplace"></a>AfxStrReplace

Replaces all the occurrences of *wszMatchStr* in *wszMainstr* with the contents of *wszReplaceWith*. Case sensitive.

```
FUNCTION AfxStrReplace (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING, BYREF wszReplaceWith AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string from which you want to replace the specified string. |
| *wszMatchStr* | The string expression to be replaced. |
| *wszReplaceWith* | The replacement string. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrReplace("Hello World", "World", "Earth")   ' Returns "Hello Earth"
```

# <a name="AfxStrReplaceI"></a>AfxStrReplaceI

Replaces all the occurrences of *wszMatchStr* in *wszMainstr* with the contents of *wszReplaceWith*. Case insensitive.

```
FUNCTION AfxStrReplaceI (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING, BYREF wszReplaceWith AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string from which you want to replace the specified string. |
| *wszMatchStr* | The string expression to be replaced. |
| *wszReplaceWith* | The replacement string. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrReplaceI("Hello world", "World", "Earth")   ' -> "Hello Earth"
```

# <a name="AfxStrReplaceAny"></a>AfxStrReplaceAny

Replaces all the occurrences of any of the individual characters *wszMatchStr* in *wszMainstr* with the contents of *wszReplaceWith*. *wszReplaceWith* must be a single character (this function does not replace words; therefore, *wszMatchStr* will not shrink or grow). Case sensitive. 

```
FUNCTION AfxStrReplaceAny  (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING, BYREF wszReplaceWith AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string from which you want to replace the specified characters. |
| *wszMatchStr* | The characters to be replaced. |
| *wszReplaceWith* | The replacement character. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrReplaceAny("abacadabra", "bac", "*")   ' -> *****d**r*
```

# <a name="AfxStrReplaceAnyI"></a>AfxStrReplaceAnyI

Replaces all the occurrences of any of the individual characters *wszMatchStr* in *wszMainstr* with the contents of *wszReplaceWith*. *wszReplaceWith* must be a single character (this function does not replace words; therefore, *wszMatchStr* will not shrink or grow). Case insensitive. 

```
FUNCTION AfxStrReplaceAnyI  (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszMatchStr AS CONST WSTRING, BYREF wszReplaceWith AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string from which you want to replace the specified characters. |
| *wszMatchStr* | The characters to be replaced. |
| *wszReplaceWith* | The replacement character. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrReplaceAnyI("abacadabra", "BaC", "*")   ' -> *****d**r*
```

# <a name="AfxStrRetain"></a>AfxStrRetain

Returns a string containing only the characters contained in a specified match string. Case sensitive.

```
FUNCTION AfxStrRetain (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string from which you want to replace the specified string. |
| *wszMatchStr* | The string expression to be searched. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRetain("abacadabra","b")   ' -> "bb"
```

# <a name="AfxStrRetainI"></a>AfxStrRetainI

Returns a string containing only the characters contained in a specified match string. Case insensitive.

```
FUNCTION AfxStrRetainI (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string from which you want to replace the specified string. |
| *wszMatchStr* | The string expression to be searched. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRetainI("abacadabra","B")   ' -> "bb"
```

# <a name="AfxStrRetainAny"></a>AfxStrRetainAny

Returns a string containing only the characters contained in a specified group of characters. Case sensitive.

```
FUNCTION AfxStrRetainAny (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string from which you want to replace the specified string. |
| *wszMatchStr* |A list of single characters to be searched for individually. A match on any one of which will cause that character to be removed from the result. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRetainAny("<p>1234567890<ak;lk;l>1234567890</p>", "<;/p>")
```

# <a name="AfxStrRetainAnyI"></a>AfxStrRetainAnyI

Returns a string containing only the characters contained in a specified group of characters. Case insensitive.

```
FUNCTION AfxStrRetainAnyI (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string from which you want to replace the specified string. |
| *wszMatchStr* |A list of single characters to be searched for individually. A match on any one of which will cause that character to be removed from the result. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRetainAnyI("<p>1234567890<ak;lk;l>1234567890</p>", "<;/P>")
```

# <a name="AfxStrReverse"></a>AfxStrReverse

Reverses the contents of a string expression.

```
FUNCTION AfxStrReverse (BYREF wszMainStr AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The string to be reversed. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrReverse("garden")   ' Returns "nedrag"
```

# <a name="AfxStrRSet"></a>AfxStrRSet

Returns a string containing a right justified string.

```
FUNCTION AfxStrRSet (BYREF wszMainStr AS CONST WSTRING, _
   BYVAL nStringLength AS LONG, BYREF wszPadCharacter AS CONST WSTRING = " ") AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The string to be justified. |
| *nStringLength* | The length of the new string. |
| *wszPadCharacter* | The character to be used for padding. If it is not specified, the string will be padded with spaces. |

#### Usage example

```
DIM cws AS CWSTR = AfxStrRSet("FreeBasic", 20, "*")
```

# <a name="AfxStrShrink"></a>AfxStrShrink

Shrinks a string to use a consistent single character delimiter.

```
FUNCTION AfxStrShrink (BYREF wszMainStr AS CONST WSTRING, BYREF wszMask AS CONST WSTRING = " ") AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMask* | One or more character delimiters to shrink. |

#### Remarks

This function creates a string with consecutive words separated by a consistent single character, making it easier to parse. If *wszMask* is not specified, all leading and trailing spaces are removed and all occurrences of two or more spaces are changed to a single space. If *wszMask* contains one or more characters to shrink, all the leading and trailing occurences of them are removes and all occurrences of one or more characters of the mask are replaced with the first character of the mask.

#### Usage example

```
DIM cws AS CWSTR = AfxStrShrink(",,, one , two     three, four,", " ,")  ' Returns "one two three four"
```

# <a name="AfxStrSplit"></a>AfxStrSplit

Splits a string into tokens, which are sequences of contiguous characters separated by any of the characters that are part of delimiters.

```
FUNCTION AfxStrSplit (BYREF wszStr AS CONST WSTRING, BYREF wszDelimiters AS WSTRING = " ") AS CSafeArray
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszStr* | The string to split. |
| *wszDelimiters* | The delimiter characters. |

#### Return value

A CSafeArray containing a token in each element.

#### Usage example

```
DIM cws AS CWSTR = "- This, a sample string."
DIM cwsa AS CSafeArray = AfxStrSplit(cws, " ,.-")
FOR i AS LONG = cwsa.LBound TO cwsa.UBound
  PRINT cwsa.GetStr(i)
NEXT
```
**Include file**: CSafeArray.inc

# <a name="AfxStrSpn"></a>AfxStrSpn

Returns the length of the initial portion of a string which consists only of characters that are part of a specified set of characters.

```
FUNCTION AfxStrSpn (BYREF wszText AS CONST WSTRING, BYREF wszSet AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | The string to be scanned. |
| *wszSet* | The set of characters for which to search. |

#### Usage example

```
DIM wszText AS WSTRING * 260 = "129th"
DIM wszSet AS WSTRING * 260 = "1234567890"
DIM n AS LONG = StrSpnW(@wszText, @wszSet)
printf(!"The initial number has %d digits.\n", n)
```

# <a name="AfxStrTally"></a>AfxStrTally

Count the number of occurrences of a string or a list of characters within a string. Case sensitive.

```
FUNCTION AfxStrTally (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string expression to be searched. |


#### Return value

The number of occurrences of *wszMatchStr* in *wszMainStr*.

#### Usage example

```
DIM nCount AS LONG = AfxStrTally("abacadabra", "ab")   ' Returns 2
```

# <a name="AfxStrTallyI"></a>AfxStrTallyI

Count the number of occurrences of a string or a list of characters within a string. Case insensitive.

```
FUNCTION AfxStrTallyI (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string expression to be searched. |


#### Return value

The number of occurrences of *wszMatchStr* in *wszMainStr*.

#### Usage example

```
DIM nCount AS LONG = AfxStrTallyI("abacadabra", "Ab")   ' -> 2
```

# <a name="AfxStrTallyAny"></a>AfxStrTallyAny

Count the number of occurrences of a string or a list of characters within a string. Case sensitive.

```
FUNCTION AfxStrTallyAny (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | A list of single characters to be searched for individually. A match on any one of which will cause the count to be incremented for each occurrence of that character. Note that repeated characters in *wszMatchStr* will not increase the count. |


#### Return value

The number of occurrences of *wszMatchStr* in *wszMainStr*.

#### Usage example

```
DIM nCount AS LONG = AfxStrTallyAny("abacadabra", "bac")   ' -> 8
```

# <a name="AfxStrTallyAnyI"></a>AfxStrTallyAnyI

Count the number of occurrences of a string or a list of characters within a string. Case insensitive.

```
FUNCTION AfxStrTallyAnyI (BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | A list of single characters to be searched for individually. A match on any one of which will cause the count to be incremented for each occurrence of that character. Note that repeated characters in *wszMatchStr* will not increase the count. |

#### Return value

The number of occurrences of *wszMatchStr* in *wszMainStr*.

#### Usage example

```
DIM nCount AS LONG = AfxStrTallyAnyI("abacadabra", "bAc")   ' -> 8
```

# <a name="AfxStrUCase"></a>AfxStrUCase

Returns an uppercased version of a string.

```
FUNCTION AfxStrUCase (BYVAL pwszStr AS WSTRING PTR, _
   BYVAL pwszLocaleName AS WSTRING PTR = LOCALE_NAME_USER_DEFAULT, _
   BYVAL dwMapFlags AS DWORD = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszStr* | The main string. |
| *pwszLocaleName* | Optional. Pointer to a locale name or one of these pre-defined values: LOCALE_NAME_INVARIANT, LOCALE_NAME_SYSTEM_DEFAULT, LOCALE_NAME_USER_DEFAULT |
| *dwMapFlags* | Optional. Flag specifying the type of transformation to use during string mapping or the type of sort key to generate. |

For a table of language culture names see: [Table of Language Culture Names, Codes, and ISO Values](https://docs.microsoft.com/en-us/previous-versions/commerce-server/ee825488(v=cs.20))

For a complete list see: [LCMapStringEx function](https://docs.microsoft.com/en-us/windows/desktop/api/winnls/nf-winnls-lcmapstringex)

#### Remarks

The string conversion functions available in FreeBasic are not fully suitable for some languages. For example, the Turkish word "karışıklığı" is uppercased as "KARıŞıKLıĞı" instead of "KARIŞIKLIĞI", and "KARIŞIKLIĞI" is lowercased to "karişikliği" instead of "karışıklığı". Notice the "ı", that is not an "i".

#### Return value

The uppercased string.

# <a name="AfxStrUnWrap"></a>AfxStrUnWrap

Removes paired characters to the beginning and end of a string.

```
FUNCTION AfxStrUnWrap (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszLeftChar AS CONST WSTRING, BYREF wszRightChar AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszLeftChar* | The left character. |
| *wszRightChar* | The right character. |

```
FUNCTION AfxStrUnWrap (BYREF wszMainStr AS CONST WSTRING, BYREF wszChar AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszChar* | The same character for both the left and right sides. |

#### Remarks

If only one wrap character/string is specified then that character or string is used for both sides. If no wrap character/string is specified then double quotes are used.

#### Usage examples

```
AfxStrUnWrap("<Paul>", "<", ">") results in Paul
AfxStrUnWrap("'Paul'", "'") results in Paul
AfxStrUnWrap("""Paul""") results in Paul
```

# <a name="AfxStrVerify"></a>AfxStrVerify

Determine whether each character of a string is present in another string. Case sensitive.

```
FUNCTION AfxStrVerify (BYVAL nStart AS LONG, _
   BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string expression to be searched. |

#### Return value

Returns zero if each character in wszMainStr is present in *wszMatchStr*; otherwise, it returns the position of the first non-matching character in *wszMainStr*. If *nStart* is zero, a negative number of a value greater that the length of wszMainstr, zero is returned.

#### Usage example

```
DIM nCount AS LONG = AfxStrVerify(5, "123.65,22.5", "0123456789")   ' Returns 7
```

#### Remark

Returns 7 since 5 starts it past the first non-digit "." at position 4.

# <a name="AfxStrVerifyI"></a>AfxStrVerifyI

Determine whether each character of a string is present in another string. Case insensitive.

```
FUNCTION AfxStrVerifyI (BYVAL nStart AS LONG, _
   BYREF wszMainStr AS CONST WSTRING, BYREF wszMatchStr AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszMatchStr* | The string expression to be searched. |

#### Return value

Returns zero if each character in wszMainStr is present in *wszMatchStr*; otherwise, it returns the position of the first non-matching character in *wszMainStr*. If *nStart* is zero, a negative number of a value greater that the length of wszMainstr, zero is returned.

#### Usage example

```
DIM nCount AS LONG = AfxStrVerifyI(5, "123.65abcx22.5", "0123456789ABC")   ' -> 10
```

# <a name="AfxStrWrap"></a>AfxStrWrap

Adds paired characters to the beginning and end of a string.

```
FUNCTION AfxStrWrap (BYREF wszMainStr AS CONST WSTRING, _
   BYREF wszLeftChar AS CONST WSTRING, BYREF wszRightChar AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszLeftChar* | The left character. |
| *wszRightChar* | The right character. |

```
FUNCTION AfxStrWrap (BYREF wszMainStr AS CONST WSTRING, BYREF wszChar AS CONST WSTRING) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszMainStr* | The main string. |
| *wszChar* | The same character for both the left and right sides. |

#### Remarks

If only one wrap character/string is specified then that character or string is used for both sides. If no wrap character/string is specified then double quotes are used.

#### Usage examples

```
AfxStrWrap("Paul", "<", ">") results in <Paul>
AfxStrWrap("Paul", "'") results in 'Paul'
AfxStrWrap("Paul") results in "Paul"
```

# <a name="AfxBase64DecodeA"></a>AfxBase64DecodeA

Converts the contents of a Base64 mime encoded string to an ascii string.

```
FUNCTION AfxBase64DecodeA (BYREF strData AS STRING) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *strData* | The string to decode. |

#### Return value

The decoded string on success, or a null string on failure.

#### Remaks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

If we want to encode a unicode string, we must convert it to utf8 before calling AfxBase64EncodeA, e.g.

````
DIM cws AS CWSTR = "おはようございます – Good morning!"
DIM s AS STRING = AfxBase64EncodeA(cws.Utf8)
````

To decode it, we can use

````
DIM cwsOut AS CWSTR = CWSTR(AfxBase64DecodeA(s), CP_UTF8)
````

or

````
DIM cwsOut AS CWSTR
cws.utf8 = AfxBase64DecodeA(s)
````

# <a name="AfxBase64DecodeW"></a>AfxBase64DecodeW

Converts the contents of a Base64 mime encoded string to an unicode string.

```
FUNCTION AfxBase64DecodeW (BYREF cwsData AS CWSTR) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsData* | The string to decode. |

#### Return value

The decoded string on success, or a null string on failure.

#### Remaks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

# <a name="AfxBase64EncodeA"></a>AfxBase64EncodeA

Converts the contents of an ascii string to Base64 mime encoding.

```
FUNCTION AfxBase64EncodeA (BYREF strData AS STRING) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *strData* | The string to encode. |

#### Return value

The encoded string on succeess, or a null string on failure.

#### Remarks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

If we want to encode a unicode string, we must convert it to utf8 before calling AfxBase64Encode, e.g.

````
DIM cws AS CWSTR = "おはようございます – Good morning!"
DIM s AS STRING = AfxBase64EncodeA(cws.Utf8)
````

To decode it, we can use

````
DIM cwsOut AS CWSTR = CWSTR(AfxBase64DecodeA(s), CP_UTF8)
````

or

````
DIM cwsOut AS CWSTR
cws.utf8 = AfxBase64DecodeA(s)
````

# <a name="AfxBase64EncodeW"></a>AfxBase64EncodeW

Converts the contents of an unicode string to Base64 mime encoding.

```
FUNCTION AfxBase64EncodWeA (BYREF cwsData AS CWSTR) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cwsData* | The string to encode. |

#### Return value

The encoded string on succeess, or a null string on failure.

#### Remarks

Base64 is a group of similar encoding schemes that represent binary data in an ASCII string format by translating it into a radix-64 representation. The Base64 term originates from a specific MIME content transfer encoding.

Base64 encoding schemes are commonly used when there is a need to encode binary data that needs be stored and transferred over media that are designed to deal with textual data. This is to ensure that the data remains intact without modification during transport. Base64 is used commonly in a number of applications including email via MIME, and storing complex data in XML.

# <a name="AfxCryptBinaryToString"></a>AfxCryptBinaryToString

Converts an array of bytes into a formatted string.

```
FUNCTION FUNCTION AfxCryptBinaryToStringA ( _
   BYVAL pbBinary AS CONST UBYTE PTR, _
   BYVAL cbBinary AS DWORD, _
   BYVAL dwFlags AS DWORD, _
   BYVAL pszString AS LPSTR, _
   BYVAL pcchString AS DWORD PTR _
) AS WINBOOL
```

```
FUNCTION AfxCryptBinaryToStringW ( _
   BYVAL pbBinary AS CONST UBYTE PTR, _
   BYVAL cbBinary AS DWORD, _
   BYVAL dwFlags AS DWORD, _
   BYVAL pszString AS LPWSTR, _
   BYVAL pcchString AS DWORD PTR _
) AS WINBOOL
```

| Parameter  | Description |
| ---------- | ----------- |
| *pbBinary* | A pointer to the array of bytes to be converted into a string. |
| *cbBinary* | The number of elements in the *pbBinary* array. |
| *dwFlags* | Specifies the format of the resulting formatted string. |
| *pszString* | A pointer to a buffer that receives the converted string. To calculate the number of characters that must be allocated to hold the returned string, set this parameter to NULL. The function will place the required number of characters, including the terminating NULL character, in the value pointed to by *pcchString*. |
| *pcchString* | A pointer to a DWORD variable that contains the size, in characters, of the *pszString* buffer. If *pszString* is NULL, the function calculates the length of the return string (including the terminating null character) in characters and returns it in this parameter. If *pszString* is not NULL and big enough, the function converts the binary data into a specified string format including the terminating null character, but *pcchString* receives the length in characters, not including the terminating null character. |

Values available for the *dwFlags* parameter:

| Value      | Meaning |
| ---------- | ----------- |
| CRYPT_STRING_BASE64HEADER | Base64, with certificate beginning and ending headers. |
| CRYPT_STRING_BASE64 | Base64, without headers. |
| CRYPT_STRING_BINARY | Pure binary copy. |
| CRYPT_STRING_BASE64REQUESTHEADER | Base64, with request beginning and ending headers. |
| CRYPT_STRING_HEX | Hexadecimal only. |
| CRYPT_STRING_HEXASCII | Hexadecimal, with ASCII character display. |
| CRYPT_STRING_BASE64X509CRLHEADER | Base64, with X.509 CRL beginning and ending headers. |
| CRYPT_STRING_HEXADDR | Hexadecimal, with address display. |
| CRYPT_STRING_HEXASCIIADDR | Hexadecimal, with ASCII character and address display. |
| CRYPT_STRING_HEXRAW | A raw hexadecimal string. Not supported in Windows Server 2003 and Windows XP. |
| CRYPT_STRING_STRICT | Enforce strict decoding of ASN.1 text formats. Some ASN.1 binary BLOBS can have the first few bytes of the BLOB incorrectly interpreted as Base64 text. In this case, the rest of the text is ignored. Use this flag to enforce complete decoding of the BLOB. Not suported in Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP. |

In addition to the values above, one or more of the following values can be specified to modify the behavior of the function.

| Value      | Meaning |
| ---------- | ----------- |
| CRYPT_STRING_NOCRLF | Do not append any new line characters to the encoded string. The default behavior is to use a carriage return/line feed (CR/LF) pair (0x0D/0x0A) to represent a new line. Not supported in Windows Server 2003 and Windows XP. |
| CRYPT_STRING_NOCR | Only use the line feed (LF) character (0x0A) for a new line. The default behavior is to use a CR/LF pair (0x0D/0x0A) to represent a new line. |


#### Return value

If the function succeeds, the function returns nonzero (CTRUE).
If the function fails, it returns zero (FALSE).

#### Remarks

With the exception of when **CRYPT_STRING_BINARY** encoding is used, all strings are appended with a new line sequence. By default, the new line sequence is a CR/LF pair (0x0D/0x0A). If the *dwFlags* parameter contains the **CRYPT_STRING_NOCR** flag, then the new line sequence is a LF character (0x0A). If the *dwFlags* parameter contains the **CRYPT_STRING_NOCRLF** flag, then no new line sequence is appended to the string.

# <a name="AfxCryptStringToBinary"></a>AfxCryptStringToBinary

Converts a formatted string into an array of bytes.

```
FUNCTION AfxCryptStringToBinaryA ( _
   BYVAL pszString AS LPCSTR, _
   BYVAL cchString AS DWORD, _
   BYVAL dwFlags AS DWORD, _
   BYVAL pbBinary AS UBYTE PTR, _
   BYVAL pcbBinary AS DWORD PTR, _
   BYVAL pdwSkip AS DWORD PTR, _
   BYVAL pdwFlags AS DWORD PTR _
) AS WINBOOL
```

```
FUNCTION AfxCryptStringToBinaryW ( _
   BYVAL pszString AS LPCWSTR, _
   BYVAL cchString AS DWORD, _
   BYVAL dwFlags AS DWORD, _
   BYVAL pbBinary AS UBYTE PTR, _
   BYVAL pcbBinary AS DWORD PTR, _
   BYVAL pdwSkip AS DWORD PTR, _
   BYVAL pdwFlags AS DWORD PTR _
) AS WINBOOL
```

| Parameter  | Description |
| ---------- | ----------- |
| *pszString* | A pointer to a string that contains the formatted string to be converted. |
| *cchString* | The number of characters of the formatted string to be converted, not including the terminating NULL character. If this parameter is zero, pszString is considered to be a null-terminated string. |
| *dwFlags* | Indicates the format of the string to be converted. |
| *pbBinary* | A pointer to a buffer that receives the returned sequence of bytes. If this parameter is NULL, the function calculates the length of the buffer needed and returns the size, in bytes, of required memory in the DWORD pointed to by *pcbBinary*. |
| *pcbBinary* | A pointer to a DWORD variable that, on entry, contains the size, in bytes, of the *pbBinary* buffer. After the function returns, this variable contains the number of bytes copied to the buffer. If this value is not large enough to contain all of the data, the function fails and GetLastError returns **ERROR_MORE_DATA**. If *pbBinary* is NULL, the DWORD pointed to by *pcbBinary* is ignored. |
| *pdwSkip* | A pointer to a DWORD value that receives the number of characters skipped to reach the beginning of the actual base64 or hexadecimal strings. This parameter is optional and can be NULL if it is not needed. |
| *pdwFlags* | A pointer to a DWORD value that receives the flags actually used in the conversion. These are the same flags used for the *dwFlags* parameter. In many cases, these will be the same flags that were passed in the *dwFlags* parameter. If *dwFlags* contains one of the flags inicated below, this value will receive a flag that indicates the actual format of the string. This parameter is optional and can be NULL if it is not needed. |

Values available for the *dwFlags* parameter:

| Value      | Meaning |
| ---------- | ----------- |
| CRYPT_STRING_BASE64HEADER | Base64, with certificate beginning and ending headers. |
| CRYPT_STRING_BASE64 | Base64, without headers. |
| CRYPT_STRING_BINARY | Pure binary copy. |
| CRYPT_STRING_BASE64REQUESTHEADER | Base64, with request beginning and ending headers. |
| CRYPT_STRING_BASE64 | Base64, without headers. |
| CRYPT_STRING_HEX | Hexadecimal only. |
| CRYPT_STRING_HEXASCII | Hexadecimal, with ASCII character display. |
| CRYPT_STRING_BASE64_ANY | Tries the following, in order: CRYPT_STRING_BASE64HEADER, CRYPT_STRING_BASE64. |
| CRYPT_STRING_ANY | Tries the following, in order: CRYPT_STRING_BASE64HEADER, CRYPT_STRING_BASE64, CRYPT_STRING_BINARY. |
| CRYPT_STRING_HEX_ANY | Tries the following, in order: CRYPT_STRING_HEXADDR, CRYPT_STRING_HEXASCIIADDR, CRYPT_STRING_HEX, CRYPT_STRING_HEXRAW, CRYPT_STRING_HEXASCII. |
| CRYPT_STRING_HEX | Hexadecimal only. |
| CRYPT_STRING_BASE64X509CRLHEADER | Base64, with X.509 CRL beginning and ending headers. |
| CRYPT_STRING_HEXADDR | Hexadecimal, with address display. |
| CRYPT_STRING_HEXASCIIADDR | Hexadecimal, with ASCII character and address display. |
| CRYPT_STRING_HEXRAW | A raw hexadecimal string. Not supported in Windows Server 2003 and Windows XP. |
| CRYPT_STRING_STRICT | Enforce strict decoding of ASN.1 text formats. Some ASN.1 binary BLOBS can have the first few bytes of the BLOB incorrectly interpreted as Base64 text. In this case, the rest of the text is ignored. Use this flag to enforce complete decoding of the BLOB. Not suported in Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP. |

Values available for the *pdwFlags* parameter:

| Value      | Meaning |
| ---------- | ----------- |
| CRYPT_STRING_ANY | This variable will receive one of the following values. Each value indicates the actual format of the string. CRYPT_STRING_BASE64HEADER, CRYPT_STRING_BASE64, CRYPT_STRING_BINARY. |
| CRYPT_STRING_BASE64_ANY | This variable will receive one of the following values. Each value indicates the actual format of the string. CRYPT_STRING_BASE64HEADER, CRYPT_STRING_BASE64. |
| CRYPT_STRING_HEX_ANY | This variable will receive one of the following values. Each value indicates the actual format of the string. CRYPT_STRING_HEXADDR, CRYPT_STRING_HEXASCIIADDR, CRYPT_STRING_HEX, CRYPT_STRING_HEXRAW, CRYPT_STRING_HEXASCII. |

