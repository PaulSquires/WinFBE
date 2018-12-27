# CRegExp Class

**CRegExp** is a wrapper class on top of the VB Script Regular Expressions.

In a typical search and replace operation, you must provide the exact text for which you are searching. That technique may be adequate for simple search and replace tasks in static text, but it lacks flexibility and makes searching dynamic text difficult, if not impossible.

With regular expressions, you can:

* Test for a pattern within a string. For example, you can test an input string to see if a telephone number pattern or a credit card number pattern occurs within the string. This is called data validation.

* Replace text. You can use a regular expression to identify specific text in a document and either remove it completely or replace it with other text.

* Extract a substring from a string based upon a pattern match. You can find specific text within a document or input field.

**Include file**: CRegExp.inc

**Tutorial**: [Introduction to Regular Expressions](#Introduction)

**Recipes**: [Recipes](#Recipes)

### Constructors

```
CONSTRUCTOR CRegExp (BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE)
```
```
CONSTRUCTOR CRegExp (BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsPattern* | The regular expression pattern being searched for. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

```
CONSTRUCTOR CRegExp (BYREF pRegExp AS CRegExp)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pRegExp* | Reference to an instance of a **CRegExp** class to be cloned. |

### Methods

| Name  | Description |
| ---------- | ----------- |
| [Execute](#Execute) | Executes a regular expression search against a specified string. |
| [Extract](#Extract) | Extracts a substring using VBScript regular expressions search patterns. |
| [Find](#Find) | Find function with VBScript regular expressions search patterns. |
| [FindEx](#FindEx) | Global, multiline find function with VBScript regular expressions search patterns. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [MatchCount](#MatchCount) | Returns the number of matches found. |
| [RegExpPtr](#RegExpPtr) | Returns a direct pointer to the **Afx_IRegExp2** interface. |
| [Remove](#Remove) | Returns a copy of a string with text removed using a regular expression as the search string. |
| [Replace](#Replace) | Replaces text found in a regular expression search. |
| [SubMatchValue](#SubMatchValue) | Retrieves the content of the specified submatch. |
| [Test](#Test) | Executes a regular expression search against a specified string and returns a boolean value that indicates if a pattern match was found. |

### Properties

| Name  | Description |
| ---------- | ----------- |
| [Global](#Global) | Sets or returns a boolean value that indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| [IgnoreCase](#IgnoreCase) | Sets or returns a boolean value that indicates if a pattern search is case-sensitive or not. |
| [MatchLen](#MatchLen) | Returns the length of a match found in a search string. |
| [MatchPos](#MatchPos) | Returns the position in a search string where a match occurs. |
| [MatchValue](#MatchValue) | Returns the value or text of a match found in a search string. |
| [Multiline](#Multiline) | Sets or returns a boolean value that indicates whether or not to search in strings across multiple lines. |
| [Pattern](#Pattern) | Sets or returns a boolean value that indicates whether or not to search in strings across multiple lines. |
| [SubMatchesCount](#SubMatchesCount) | Returns the number of submatches. |

# <a name="Execute"></a>Execute

Executes a regular expression search against a specified string.

```
FUNCTION Execute (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
```
```
FUNCTION Execute (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYREF cvReplaceString AS CVAR,BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The main string. |
| *cbsPattern* | The regular string expression being searched for. |
| *cvReplaceString* | The replacement text string. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not.  |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Remarks

In the first overloaded method, the actual pattern for the regular expression search is set using the **Pattern** property.

#### Return value

BOOLEAN. True on success or False on failure.

#### Example

```
'#CONSOLE ON
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
pRegExp.Pattern = "(\w+)@(\w+)\.(\w+)"
pRegExp.IgnoreCase = TRUE
DIM cbsText AS CBSTR = "Please send mail to dragon@xyzzy.com. Thanks!"
IF pRegExp.Execute(cbsText) = FALSE THEN
   print "No match found"
ELSE
   ' Get the number of submatches
   DIM nCount AS LONG = pRegExp.SubMatchesCount(0)
   print "Sub matches: ", nCount
   FOR i AS LONG = 0 TO nCount - 1
      print pRegExp.SubMatchValue(0, i)
   NEXT
END IF

PRINT
PRINT "Press any key..."
SLEEP
```

# <a name="Extract"></a>Extract

Extracts a substring using VBScript regular expressions search patterns.

```
FUNCTION Extract (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Extract (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = FALSE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Extract (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, _
   BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The position in the string at which the search will begin. The first character starts at position 1. |
| *cbsSourceString* | The text to be parsed. |
| *cbsPattern* | The pattern to match. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Return value

CBSTR. The retrieved string.

#### Usage examples

```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbs AS CBSTR = pRegExp.Extract(cbsText, cbsPattern)
' Output: a234
```
```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbs AS CBSTR = pRegExp.Extract(15, cbsText, cbsPattern)
' Output: x345
```
```
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbs AS CBSTR = CRegExp(cbsPattern).Extract(cbsText)
' Output: a234
```
```
' // Ignore case
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsText AS CBSTR = "blah blah A234 blah blah x345 blah blah"
DIM cbs AS CBSTR = CRegExp(cbsPattern).Extract(cbsText, TRUE)
' Output: A234
```

# <a name="Find"></a>Find

Find function with VBScript regular expressions search patterns.

```
FUNCTION Find (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
```
```
FUNCTION Find (BYREF cbsSourceString AS CBSTR, _
   BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
```
```
FUNCTION Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
```
```
FUNCTION Find (BYVAL nStart AS LONG, BYREF cbsSourceString AS CBSTR, _
   BYREF cbsPattern AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStart* | The position in the string at which the search will begin. The first character starts at position 1. |
| *cbsSourceString* | The text to be parsed. |
| *cbsPattern* | The pattern to match. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Return value

Returns the position of the match or 0 if not found. The length of the match can be retrieved calling the **MatchLen** property.

#### Usage examples

```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM nPos AS LONG = pRegExp.Find(cbsText, cbsPattern)
' Output: 11
```
```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM nPos AS LONG = pRegExp.Find(15, cbsText, cbsPattern)
' Output: 26
```
```
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM nPos AS LONG = CRegExp(cbsPattern).Find(cbsText)
' Output: 11
```
```
DIM nPos AS LONG = CRegExp("[a-z][0-9][0-9][0-9]").Find("blah blah a234 blah blah x345 blah blah")
' Output: 11
```

# <a name="FindEx"></a>FindEx

Global, multiline find function with VBScript regular expressions search patterns.

```
FUNCTION FindEx (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = TRUE) AS CBSTR
```
```
FUNCTION FindEx (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, _
   BYVAL bMultiline AS BOOLEAN = TRUE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The text to be parsed. |
| *cbsPattern* | The pattern to match. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Return value

Returns a list of comma separated "index, length" value pairs. The pairs are separated by a semicolon.

#### Usage example

```
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsOut AS CBSTR = pRegExp.FindEx(cbsText, cbsPattern)
' Output: 11,4;26,4
```
```
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsOut AS CBSTR = CRegExp(cbsPattern).FindEx(cbsText)
' Output: 11,4;26,4
```
# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```
#### Return value

S_OK (0) on success, or an error code on failure.

# <a name="MatchCount"></a>MatchCount

Returns the number of matches found.

```
FUNCTION MatchCount () AS LONG
```

# <a name="RegExpPtr"></a>RegExpPtr

Returns a direct pointer to the Afx_IRegExp2 interface.

```
FUNCTION RegExpPtr () AS Afx_IRegExp2 PTR
```
#### Remarks

Since it is a direct pointer, you don't have to release it calling the **Release** method.

# <a name="Remove"></a>Remove

Returns a copy of a string with text removed using a regular expression as the search string.

```
FUNCTION Remove (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Remove (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The main string. |
| *cbsPattern* | The pattern to be removed. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Usage examples

```
DIM pRegExp AS CRegExp
PRINT pRegExp.Remove("abacadabra", "ab") ' - prints "acadra"
PRINT pRegExp.Remove("abacadabra", "[bAc]", TRUE) ' - prints "dr"
PRINT pRegExp.Remove("World, worldx, world", $"\bworld\b", TRUE) ' prints ", worldx,"
```
```
PRINT CRegExp("ab").Remove("abacadabra") ' - prints "acadra"
PRINT CRegExp("[bAc]").Remove("abacadabra", TRUE) ' - prints "dr"
PRINT CRegExp($"\bworld\b").Remove("World, worldx, world", TRUE) ' prints ", worldx,"
```

# <a name="Replace"></a>Replace

Replaces text found in a regular expression search.

```
FUNCTION Replace (BYREF cbsSourceString AS CBSTR, BYREF cvReplaceString AS CVAR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bGlobal AS BOOLEAN = TRUE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```
```
FUNCTION Replace (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYREF cvReplaceString AS CVAR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bGlobal AS BOOLEAN = TRUE, BYVAL bMultiline AS BOOLEAN = FALSE) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The main string. |
| *cvReplaceString* | The replacement text string. |
| *cbsPattern* | The regular string expression being searched for. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bGlobal* | TRUE or FALSE. Indicates if a pattern should match all occurrences in an entire search string or just the first one. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Remarks

In the first overloaded function, the actual pattern for the text being replaced is set using the Pattern property.

The **Replace** method returns a copy of *cbsSourceString* with the text of *cbsPattern* replaced with *cvsReplaceString*. If no match is found, a copy of *cbsSourceString* is returned unchanged.

#### Examples

```
'#CONSOLE ON
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
pRegExp.Pattern = "fox"
pRegExp.IgnoreCase = TRUE
DIM cbsText AS CBSTR = "The quick brown fox jumped over the lazy dog."
' Make replacement
DIM cbsRes AS CBSTR = pRegExp.Replace(cbsText, "cat")
print cbsRes
' Output: The quick brown cat jumped over the lazy dog.
```
In addition, the **Replace** method can replace subexpressions in the pattern.

The following call to the function shown in the previous example swaps the first pair of words in the original string:

```
DIM pRegExp AS CRegExp
pRegExp.Pattern = "(\S+)(\s+)(\S+)"
pRegExp.IgnoreCase = TRUE
DIM cbsText AS CBSTR = "The quick brown fox jumped over the lazy dog."
' Make replacement
DIM cbsRes AS CBSTR = pRegExp.Replace(cbsText, "$3$2$1")
print cbsRes
```

Suppose that you have a telephone directory, and all the phone numbers are formatted like this:
555-123-4567. Now, you decide that all the phone numbers should be formatted to look like this: (555) 123-4567.

```
DIM pRegExp AS CRegExp
pRegExp.Global = TRUE
pRegExp.Pattern = "(\d{3})-(\d{3})-(\d{4})"
DIM cbsText AS CBSTR = "555-123-4567"
DIM cbsRes AS CBSTR = pRegExp.Replace(cbsText, "($1) $2-$3")
print cbsRes
```
--or--

```
DIM cbsText AS CBSTR = "555-123-4567"
DIM cbsPattern AS CBSTR = "(\d{3})-(\d{3})-(\d{4})"
DIM cbsRes AS CBSTR = CRegExp(cbsPattern).Replace(cbsText, "($1) $2-$3")
print cbsRes
```

What we have done is to search for 3 digits (\d{3}) followed by a dash, followed by 3 more digits and a dash, followed by 4 digits and add () to the first three digits and change the first dash with a space.  $1, $2, and $3 are examples of a regular expression "back reference." A back reference is simply a portion of the found text that can be saved and then reused.

# <a name="SubMatchValue"></a>SubMatchValue

Retrieves the content of the specified submatch.

```
FUNCTION SubMatchValue (BYVAL MatchIndex AS LONG = 0, BYVAL SubMatchIndex AS LONG = 0) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *MatchIndex* | 0-based index of the match to retrieve. |
| *SubMatchIndex* | 0-based index of the submatch to retrieve. |

# <a name="Test"></a>Test

Executes a regular expression search against a specified string and returns a boolean value that indicates if a pattern match was found.

```
FUNCTION Test (BYREF cbsSourceString AS CBSTR, BYVAL bIgnoreCase AS BOOLEAN = FALSE, _
   BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
```
```
FUNCTION Test (BYREF cbsSourceString AS CBSTR, BYREF cbsPattern AS CBSTR, _
   BYVAL bIgnoreCase AS BOOLEAN = FALSE, BYVAL bMultiline AS BOOLEAN = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsSourceString* | The main string. |
| *cbsPattern* | The regular string expression being searched for. |
| *bIgnoreCase* | TRUE or FALSE. Indicates if a pattern search is case-sensitive or not. |
| *bMultiline* | TRUE or FALSE. Whether or not to search in strings across multiple lines. |

#### Return value

BOOLEAN. True if a pattern match is found; False if no match is found.

#### Remarks

In the first overloaded method, the actual pattern for the regular expression search is set using the **Pattern** property. The **Global** property has no effect on the **Test** method.

# <a name="Global"></a>Global

Sets or returns a boolean value that indicates if a pattern should match all occurrences in an entire search string or just the first one.

```
PROPERTY Global () AS BOOLEAN
PROPERTY Global (BYVAL bGlobal AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bGlobal* | True if the search applies to the entire string, False if it does not. Default is False. |

# <a name="IgnoreCase"></a>IgnoreCase

Sets or returns a boolean value that indicates if a pattern search is case-sensitive or not.

```
PROPERTY IgnoreCase () AS BOOLEAN
PROPERTY IgnoreCase (BYVAL bIgnoreCase AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bIgnoreCase* | False if the search is case-sensitive, True if it is not. Default is False. |

# <a name="MatchLen"></a>MatchLen

Returns the length of a match found in a search string.

```
PROPERTY MatchLen (BYVAL index AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

# <a name="MatchPos"></a>MatchPos

Returns the position in a search string where a match occurs.

```
PROPERTY MatchPos (BYVAL index AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

#### Remarks

The **MatchPos** property uses a zero-based offset from the beginning of the search string. In other words, the first character in the string is identified as character zero (0).

# <a name="MatchValue"></a>MatchValue

Returns the value or text of a match found in a search string.

```
PROPERTY MatchValue (BYVAL index AS LONG = 0) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

# <a name="Multiline"></a>Multiline

Sets or returns a boolean value that indicates whether or not to search in strings across multiple lines.

```
PROPERTY Multiline () AS BOOLEAN
PROPERTY Multiline (BYVAL bMultiline AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bMultiline* | True if the search is performed across multiple lines, False if it is not. Default is False. |

# <a name="Pattern"></a>Pattern

Sets or returns the regular expression pattern being searched for.

```
PROPERTY Pattern () AS CWSTR
PROPERTY Pattern (BYREF cbsPattern AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsPattern* | Regular string expression being searched for. May include any of the regular expression characters defined in the table in the Settings section. |

#### Settings

Special characters and sequences are used in writing patterns for regular expressions. The following table describes and gives an example of the characters and sequences that can be used.

| Character  | Description |
| ---------- | ----------- |
| \\ | Marks the next character as either a special character or a literal. For example, "n" matches the character "n". "\\n" matches a newline character. The sequence "\\\\" matches "\\" and "\\(" matches "(". |
| ^ | Matches the beginning of input. |
| $ | Matches the end of input. |
| * | Matches the preceding character zero or more times. For example, "zo*" matches either "z" or "zoo". |
| + | Matches the preceding character one or more times. For example, "zo+" matches "zoo" but not "z". |
| ? | Matches the preceding character zero or one time. For example, "a?ve?" matches the "ve" in "never". |
| . | Matches any single character except a newline character. |
| (pattern) | Matches pattern and remembers the match. The matched substring can be retrieved from the resulting Matches collection, using Item \[0]...\[n]. To match parentheses characters ( ), use "\\(" or "\\)". |
| x\|y | Matches either x or y. For example, "z\|wood" matches "z" or "wood". "(z\|w)oo" matches "zoo" or "wood". |
| {n} | n is a nonnegative integer. Matches exactly n times. For example, "o{2}" does not match the "o" in "Bob," but matches the first two o's in "foooood". |
| {n,} | n is a nonnegative integer. Matches at least n times. For example, "o{2,}" does not match the "o" in "Bob" and matches all the o's in "foooood." "o{1,}" is equivalent to "o+". "o{0,}" is equivalent to "o*". |
| { n , m } | m and n are nonnegative integers. Matches at least n and at most m times. For example, "o{1,3}" matches the first three o's in "fooooood." "o{0,1}" is equivalent to "o?". |
| \[xyz] | A character set. Matches any one of the enclosed characters. For example, "\[abc]" matches the "a" in "plain".  |
| \[^xyz] | A negative character set. Matches any character not enclosed. For example, "\[^abc]" matches the "p" in "plain".  |
| \[a-z] | A range of characters. Matches any character in the specified range. For example, "\[a-z]" matches any lowercase alphabetic character in the range "a" through "z". |
| \[^m-z] | A negative range characters. Matches any character not in the specified range. For example, "\[m-z]" matches any character not in the range "m" through "z".  |
| \\b | Matches a word boundary, that is, the position between a word and a space. For example, "er\b" matches the "er" in "never" but not the "er" in "verb". |
| \\B | Matches a non-word boundary. "ea\*r\B" matches the "ear" in "never early". |
| \\d | Matches a digit character. Equivalent to \[0-9]. |
| \\D | Matches a non-digit character. Equivalent to \[^0-9]. |
| \\f | Matches a form-feed character. |
| \\n | Matches a newline character. |
| \\r | Matches a carriage return character. |
| \\s | Matches any white space including space, tab, form-feed, etc. Equivalent to "\[ \f\n\r\t\v]". |
| \\S | Matches any nonwhite space character. Equivalent to "\[^ \f\n\r\t\v]".  |
| \\t | Matches a tab character. |
| \\v | Matches a vertical tab character. |
| \\w | Matches any word character including underscore. Equivalent to "\[A-Za-z0-9_]". |
| \\W | Matches any non-word character. Equivalent to "\[^A-Za-z0-9_]". |
| \\num | Matches num, where num is a positive integer. A reference back to remembered matches. For example, "(.)\1" matches two consecutive identical characters. |
| \\n | Matches n, where n is an octal escape value. Octal escape values must be 1, 2, or 3 digits long. For example, "\11" and "\011" both match a tab character. "\0011" is the equivalent of "\001" & "1". Octal escape values must not exceed 256. If they do, only the first two digits comprise the expression. Allows ASCII codes to be used in regular expressions. |
| \\xn | Matches n, where n is a hexadecimal escape value. Hexadecimal escape values must be exactly two digits long. For example, "\x41" matches "A". "\x041" is equivalent to "\x04" & "1". Allows ASCII codes to be used in regular expressions. |

#### Example

```
'#CONSOLE ON
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
pRegExp.Pattern = "is."
pRegExp.IgnoreCase = TRUE
pRegExp.Global = TRUE
IF pRegExp.Execute("IS1 is2 IS3 is4") = FALSE THEN
   print "No match found"
ELSE
   DIM nCount AS LONG = pRegExp.MatchesCount
   FOR i AS LONG = 0 TO nCount - 1
      print "Value: ", pRegExp.MatchValue(i)
      print "Position: ", pRegExp.MatchPos(i)
      print "Length: ", pRegExp.MatchLen(i)
      print
   NEXT
END IF

PRINT
PRINT "Press any key..."
SLEEP
```

# <a name="SubMatchesCount"></a>SubMatchesCount

Returns the number of submatches.

```
PROPERTY SubMatchesCount (BYVAL index AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *index* | 0-based index of the match to retrieve. |

# <a name="Introduction"></a>Introduction to Regular Expressions

Unless you have worked with regular expressions before, the term and the concept may be unfamiliar to you. However, they may not be as unfamiliar as you think.

### The Concept of Regular Expressions

Think about how you search for files on your hard disk. You most likely use the ? and * characters to find files. The ? character matches a single character in a file name, while the * matches zero or more characters. A pattern such as 'data?.dat' would find the following files:

data1.dat
data2.dat
datax.dat
dataN.dat

Using the * character instead of the ? character expands the number of files found. 'data*.dat' matches all of the following:

data.dat
data1.dat
data2.dat
data12.dat
datax.dat
dataXYZ.dat

While this method of searching for files can certainly be useful, it is also very limited. The limited ability of the ? and * wildcard characters give you an idea of what regular expressions can do, but regular expressions are much more powerful and flexible.

### Uses for Regular Expressions

In a typical search and replace operation, you must provide the exact text for which you are searching. That technique may be adequate for simple search and replace tasks in static text, but it lacks flexibility and makes searching dynamic text difficult, if not impossible. 

Flexibility of Regular Expressions

With regular expressions, you can: 

* Test for a pattern within a string. For example, you can test an input string to see if a telephone number pattern or a credit card number pattern occurs within the string. This is called data validation.

* Replace text. You can use a regular expression to identify specific text in a document and either remove it completely or replace it with other text.

* Extract a substring from a string based upon a pattern match. You can find specific text within a document or input field

For example, if you need to search an entire web site to remove some outdated material and replace some HTML formatting tags, you can use a regular expression to test each file to see if the material or the HTML formatting tags you are looking for exists in that file. That way, you can narrow down the affected files to only those that contain the material that has to be removed or changed. You can then use a regular expression to remove the outdated material, and finally, you can use regular expressions to search for and replace the tags that need replacing.

### Regular Expression Syntax

A regular expression is a pattern of text that consists of ordinary characters (for example, letters a through z) and special characters, known as metacharacters. The pattern describes one or more strings to match when searching a body of text. The regular expression serves as a template for matching a character pattern to the string being searched.

The following table contains the complete list of metacharacters and their behavior in the context of regular expressions:

| Character  | Description |
| ---------- | ----------- |
| \\ | Marks the next character as either a special character or a literal. For example, "n" matches the character "n". "\\n" matches a newline character. The sequence "\\\\" matches "\\" and "\\(" matches "(". |
| ^ | Matches the beginning of input. |
| $ | Matches the end of input. |
| * | Matches the preceding character zero or more times. For example, "zo*" matches either "z" or "zoo". |
| + | Matches the preceding character one or more times. For example, "zo+" matches "zoo" but not "z". |
| ? | Matches the preceding character zero or one time. For example, "a?ve?" matches the "ve" in "never". |
| . | Matches any single character except a newline character. |
| (pattern) | Matches pattern and remembers the match. The matched substring can be retrieved from the resulting Matches collection, using Item \[0]...\[n]. To match parentheses characters ( ), use "\\(" or "\\)". |
| x\|y | Matches either x or y. For example, "z\|wood" matches "z" or "wood". "(z\|w)oo" matches "zoo" or "wood". |
| {n} | n is a nonnegative integer. Matches exactly n times. For example, "o{2}" does not match the "o" in "Bob," but matches the first two o's in "foooood". |
| {n,} | n is a nonnegative integer. Matches at least n times. For example, "o{2,}" does not match the "o" in "Bob" and matches all the o's in "foooood." "o{1,}" is equivalent to "o+". "o{0,}" is equivalent to "o*". |
| { n , m } | m and n are nonnegative integers. Matches at least n and at most m times. For example, "o{1,3}" matches the first three o's in "fooooood." "o{0,1}" is equivalent to "o?". |
| \[xyz] | A character set. Matches any one of the enclosed characters. For example, "\[abc]" matches the "a" in "plain".  |
| \[^xyz] | A negative character set. Matches any character not enclosed. For example, "\[^abc]" matches the "p" in "plain".  |
| \[a-z] | A range of characters. Matches any character in the specified range. For example, "\[a-z]" matches any lowercase alphabetic character in the range "a" through "z". |
| \[^m-z] | A negative range characters. Matches any character not in the specified range. For example, "\[m-z]" matches any character not in the range "m" through "z".  |
| \\b | Matches a word boundary, that is, the position between a word and a space. For example, "er\b" matches the "er" in "never" but not the "er" in "verb". |
| \\B | Matches a non-word boundary. "ea\*r\B" matches the "ear" in "never early". |
| \\d | Matches a digit character. Equivalent to \[0-9]. |
| \\D | Matches a non-digit character. Equivalent to \[^0-9]. |
| \\f | Matches a form-feed character. |
| \\n | Matches a newline character. |
| \\r | Matches a carriage return character. |
| \\s | Matches any white space including space, tab, form-feed, etc. Equivalent to "\[ \f\n\r\t\v]". |
| \\S | Matches any nonwhite space character. Equivalent to "\[^ \f\n\r\t\v]".  |
| \\t | Matches a tab character. |
| \\v | Matches a vertical tab character. |
| \\w | Matches any word character including underscore. Equivalent to "\[A-Za-z0-9_]". |
| \\W | Matches any non-word character. Equivalent to "\[^A-Za-z0-9_]". |
| \\num | Matches num, where num is a positive integer. A reference back to remembered matches. For example, "(.)\1" matches two consecutive identical characters. |
| \\n | Matches n, where n is an octal escape value. Octal escape values must be 1, 2, or 3 digits long. For example, "\11" and "\011" both match a tab character. "\0011" is the equivalent of "\001" & "1". Octal escape values must not exceed 256. If they do, only the first two digits comprise the expression. Allows ASCII codes to be used in regular expressions. |
| \\xn | Matches n, where n is a hexadecimal escape value. Hexadecimal escape values must be exactly two digits long. For example, "\x41" matches "A". "\x041" is equivalent to "\x04" & "1". Allows ASCII codes to be used in regular expressions. |

Examples of Regular Expressions

"^\\s\*$" Match a blank line.<br>
"\\d{2}-\\d{5}" Validate an ID number consisting of 2 digits, a hyphen, and another 5 digits.

#### Constructing a Regular Expression

Regular expressions are constructed in the same way that arithmetic expressions are created. That is, small expressions are combined using a variety of metacharacters and operators to create larger expressions.

You construct a regular expression by putting the various components of the expression pattern between a pair of quotation marks ("") delimit regular expressions. For example: "expression". The regular expression pattern (expression) is stored in the **Pattern** property.

The components of a regular expression can be individual characters, sets of characters, ranges of characters, choices between characters, or any combination of all of these components.

Once you have constructed a regular expression, it is evaluated much like an arithmetic expression, that is, it is evaluated from left to right and follows an order of precedence. 

#### Order of Precedence

From highest to lowest, the order of precedence of the regular expression operators:

| Operator(s) | Description |
| ---------- | ----------- |
| \\ | Escape |
| (), (?:), (?=), \[] | Parentheses and brackets |
| \*, +, ?, {n}, {n,}, {n,m} | Quantifiers |
| ^, $, \\anymetacharacter | Anchors and Sequences |
| \| | Alternation |

Characters have higher precedence than the alternation operator, which allows 'm|food' to match "m" or "food". To match "mood" or "food", use parentheses to create a subexpression, which results in '(m|f)ood'.

#### Ordinary Characters

Ordinary characters consist of all printable and non-printable characters that are not explicitly designated as metacharacters. This includes all uppercase and lowercase alphabetic characters, all digits, all punctuation marks, and some symbols. 

The simplest form of a regular expression is a single, ordinary character that matches itself in a searched string. For example, the single-character pattern 'A' matches the letter 'A' wherever it appears in the searched string. Here are some examples of single-character regular expression patterns:

"a"<br>
"7"<br>
"M"

You can combine a number of single characters to form a larger expression. For example, the following regular expression is nothing more than an expression created by combining the single-character expressions 'a', '7', and 'M'. 

"a7M"

Notice that there is no concatenation operator. All that is required is that you just put one character after another.

#### Special Characters

There are a number of metacharacters that require special treatment when trying to match them. To match these special characters, you must first escape those characters, that is, precede them with a backslash character (\\). 

| Special character | Meaning |
| ---------- | ----------- |
| $ | Matches the position at the end of an input string. If the RegExp object's Multiline property is set, $ also matches the position preceding '\\n' or '\\r'. To match the $ character itself, use \\$. |
| ( ) | Marks the beginning and end of a subexpression. Subexpressions can be captured for later use. To match these characters, use \\( and \\). |
| * | Matches the preceding character or subexpression zero or more times. To match the * character, use \\\*. |
| + | Matches the preceding character or subexpression one or more times. To match the + character, use \\+. |
| . | Matches any single character except the newline character \\n. To match ., use \\. |
| \[ ] | Marks the beginning of a bracket expression. To match these characters, use \\\[ and \\]. |
| ? | Matches the preceding character or subexpression zero or one time, or indicates a non-greedy quantifier. To match the ? character, use \\?. |
| \ | Marks the next character as either a special character, a literal, a backreference, or an octal escape. For example, 'n' matches the character 'n'. '\\n' matches a newline character. The sequence '\\\\' matches "\\" and '\\(' matches "(". |
| / | Denotes the start or end of a literal regular expression. To match the '/' character, use '\\/'. |
| ^ | Matches the position at the beginning of an input string except when used in a bracket expression where it negates the character set. To match the ^ character itself, use \\^. |
| { } | Marks the beginning of a quantifier expression. To match these characters, use \\{ and \}. |
| \| | Indicates a choice between two items. To match \|, use \\\|. |

#### Escape Sequences that Represent Non-printing Characters

There are a number of useful non-printing characters that must be used occasionally.

| Character  | Meaning     |
| ---------- | ----------- |
| \\cx | Matches the control character indicated by x. For example, \\cM matches a Control-M or carriage return character. The value of x must be in the range of A-Z or a-z. If not, c is assumed to be a literal 'c' character. |
| \\f | Matches a form-feed character. Equivalent to \\x0c and \\cL. |
| \\n | Matches a newline character. Equivalent to \\x0a and \\cJ. |
| \\r | Matches a carriage return character. Equivalent to \\x0d and \\cM. |
| \\s | Matches any white space character including space, tab, form-feed, and so on. Equivalent to \[\f\n\r\t\v]. |
| \\S | Matches any non-white space character. Equivalent to \[^\f\n\r\t\v]. |
| \\t | Matches a tab character. Equivalent to \\x09 and \\cI. |
| \\v | Matches a vertical tab character. Equivalent to \\x0b and \\cK. |

#### Character Matching

The period (.) matches any single printing or non-printing character in a string, except a newline character (\n). The following regular expression matches 'aac', 'abc', 'acc', 'adc', and so on, as well as 'a1c', 'a2c', a-c', and a#c': 

"a.c"

If you are trying to match a string containing a file name where a period (.) is part of the input string, you do so by preceding the period in the regular expression with a backslash (\\) character. To illustrate, the following regular expression matches 'filename.ext':

"filename\\.ext"

These expressions are still pretty limited. They only let you match any single character. Many times, it is useful to match specified characters from a list. For example, if you have an input text that contains chapter headings that are expressed numerically as Chapter 1, Chapter 2, and so on, you might want to find those chapter headings. 

#### Bracket Expressions

You can create a list of matching characters by placing one or more individual characters within square brackets (\[ and ]). When characters are enclosed in brackets, the list is called a bracket expression. Within brackets, as anywhere else, ordinary characters represent themselves, that is, they match an occurrence of themselves in the input text. Most special characters lose their meaning when they occur inside a bracket expression. Here are some exceptions: 

The ']' character ends a list if it is not the first item. To match the ']' character in a list, place it first, immediately following the opening '\['.

The '\\' character continues to be the escape character. To match the '\\' character, use '\\\\'.

Characters enclosed in a bracket expression match only a single character for the position in the regular expression where the bracket expression appears. The following regular expression matches 'Chapter 1', 'Chapter 2', 'Chapter 3', 'Chapter 4', and 'Chapter 5':

"Chapter \[12345]"

Notice that the word 'Chapter' and the space that follows are fixed in position relative to the characters within brackets. The bracket expression then, is used to specify only the set of characters that matches the single character position immediately following the word 'Chapter' and a space. That is the ninth character position.

If you want to express the matching characters using a range instead of the characters themselves, you can separate the beginning and ending characters in the range using the hyphen (-) character. The character value of the individual characters determines their relative order within a range. The following regular expression contains a range expression that is equivalent to the bracketed list shown above.

"Chapter \[1-5]"

When a range is specified in this manner, both the starting and ending values are included in the range. It is important to note that the starting value must precede the ending value in Unicode sort order.

If you want to include the hyphen character in your bracket expression, you must do one of the following: 

Escape it with a backslash: 

\[\\-]

Put the hyphen character at the beginning or the end of the bracketed list. The following expressions matches all lowercase letters and the hyphen: 

\[-a-z]
\[a-z-]

Create a range where the beginning character value is lower than the hyphen character and the ending character value is equal to or greater than the hyphen. Both of the following regular expressions satisfy this requirement: 

\[!--]
\[!-~]

You can also find all the characters not in the list or range by placing the caret (^) character at the beginning of the list. If the caret character appears in any other position within the list, it matches itself, that is, it has no special meaning. The following regular expression matches chapter headings with numbers greater than 5':

"Chapter \[^12345]"

In the examples shown above, the expression matches any digit character in the ninth position except 1, 2, 3, 4, or 5. So, for example, 'Chapter 7' is a match and so is 'Chapter 9'. 

The same expressions above can be represented using the hyphen character (-):

"Chapter \[^1-5]"

A typical use of a bracket expression is to specify matches of any upper- or lowercase alphabetic characters or any digits. The following expression specifies such a match:

"\[A-Za-z0-9]"

#### Quantifiers

Sometimes, you do not know how many characters there are to match. In order to accommodate that kind of uncertainty, regular expressions support the concept of quantifiers. These quantifiers let you specify how many times a given component of your regular expression must occur for your match to be true.

| Character  | Meaning     |
| ---------- | ----------- |
| \* | Matches the preceding character or subexpression zero or more times. For example, 'zo*' matches "z" and "zoo". * is equivalent to {0,}. |
| + | Matches the preceding character or subexpression one or more times. For example, 'zo+' matches "zo" and "zoo", but not "z". + is equivalent to {1,}. |
| ? | Matches the preceding character or subexpression zero or one time. For example, 'do(es)?' matches the "do" in "do" or "does". ? is equivalent to {0,1}. |
| {n} | n is a nonnegative integer. Matches exactly n times. For example, 'o{2}' does not match the 'o' in "Bob," but matches the two o's in "food". |
| {n,} | n is a nonnegative integer. Matches at least n times. For example, 'o{2,}' does not match the 'o' in "Bob" and matches all the o's in "foooood". 'o{1,}' is equivalent to 'o+'. 'o{0,}' is equivalent to 'o*'. |
| {n,m} | m and n are nonnegative integers, where n <= m. Matches at least n and at most m times. For example, 'o{1,3}' matches the first three o's in "fooooood". 'o{0,1}' is equivalent to 'o?'. Note that you cannot put a space between the comma and the numbers. |

With a large input document, chapter numbers could easily exceed nine, so you need a way to handle two or three digit chapter numbers. Quantifiers give you that capability. The following regular expression matches chapter headings with any number of digits:

"Chapter \[1-9][0-9]*"

Notice that the quantifier appears after the range expression. Therefore, it applies to the entire range expression which, in this case, specifies only digits from 0 through 9, inclusive.

The '+' quantifier is not used here because there does not necessarily need to be a digit in the second or subsequent position. The '?' character also is not used because it limits the chapter numbers to only two digits. You want to match at least one digit following 'Chapter' and a space character.

If you know that your chapter numbers are limited to only 99 chapters, you can use the following expression to specify at least one, but not more than 2 digits.

"Chapter \[0-9]{1,2}"

The disadvantage to the expression shown above is that if there is a chapter number greater than 99, it will still only match the first two digits. Another disadvantage is that somebody could create a Chapter 0 and it would match. A better expression for matching only two digits are the following:
 
"Chapter \[1-9][0-9]?"

-or-
 	
"Chapter \[1-9]\[0-9]{0,1}"

The '\*', '+', and '?' quantifiers are all what are referred to as greedy, that is, they match as much text as possible. Sometimes that is not at all what you want to happen. Sometimes, you just want a minimal match. 

Say, for example, you are searching an HTML document for an occurrence of a chapter title enclosed in an H1 tag. That text appears in your document as:

\<H1>Chapter 1 – Introduction to Regular Expressions\</H1>

The following expression matches everything from the opening less than symbol (<) to the greater than symbol (>) at the end of the closing H1 tag.

"\<\.\*>"

If all you really wanted to match was the opening H1 tag, the following, non-greedy expression matches only \<H1>.

"\<\.\*?>"

By placing the '?' after a '\*', '+', or '?' quantifier, the expression is transformed from a greedy to a non-greedy, or minimal, match.

#### Anchors

So far, the examples you've seen have been concerned only with finding chapter headings wherever they occur. Any occurrence of the string 'Chapter' followed by a space, followed by a number, could be an actual chapter heading, or it could also be a cross-reference to another chapter. Since true chapter headings always appear at the beginning of a line, you'll need to devise a way to find only the headings and not find the cross-references.

Anchors provide that capability. Anchors allow you to fix a regular expression to either the beginning or end of a line. They also allow you to create regular expressions that occur either within a word or at the beginning or end of a word. The following table contains the list of regular expression anchors and their meanings:

| Character  | Meaning     |
| ---------- | ----------- |
| ^ | Matches the position at the beginning of the input string. If the RegExp object's Multiline property is set, ^ also matches the position following '\\n' or '\\r'. |
| $ | Matches the position at the end of the input string. If the RegExp object's Multiline property is set, $ also matches the position preceding '\\n' or '\\r'. |
| \\b | Matches a word boundary, that is, the position between a word and a space. |
| \\B | Matches a nonword boundary. |

You cannot use a quantifier with an anchor. Since you cannot have more than one position immediately before or after a newline or word boundary, expressions such as '^\*' are not permitted.

To match text at the beginning of a line of text, use the '^' character at the beginning of the regular expression. Do not confuse this use of the '^' with the use within a bracket expression. 

To match text at the end of a line of text, use the '$' character at the end of the regular expression.

To use anchors when searching for chapter headings, the following regular expression matches a chapter heading with up to two following digits that occurs at the beginning of a line:
 
"^Chapter \[1-9]\[0-9]{0,1}"

Not only does a true chapter heading occur at the beginning of a line, it is also the only text on the line, so it also must be at the end of a line as well. The following expression ensures that the match specified only matches chapters and not cross-references. It does so by creating a regular expression that matches only at the beginning and end of a line of text.
 
"^Chapter \[1-9]\[0-9]{0,1}$"

Matching word boundaries is a little different but adds a very important capability to regular expressions. A word boundary is the position between a word and a space. A nonword boundary is any other position. The following expression matches the first three characters of the word 'Chapter' because they appear following a word boundary:
 
"\\bCha"

The position of the '\\b' operator is critical. If it is positioned at the beginning of a string to be matched, it looks for the match at the beginning of the word; if it is positioned at the end of the string, it looks for the match at the end of the word. For example, the following expression matches 'ter' in the word 'Chapter' because it appears before a word boundary:
 
"ter\\b"

The following expression matches 'apt' as it occurs in 'Chapter', but not as it occurs in 'aptitude':
 
"\\Bapt"

The string 'apt' occurs on a nonword boundary in the word 'Chapter' but on a word boundary in the word 'aptitude'. For the \\B nonword boundary operator, position is not important because the match is not relative to the beginning or end of a word.
 
#### Alternation and Grouping

Alternation allows use of the '\|' character to allow a choice between two or more alternatives. Expanding the chapter heading regular expression, you can expand it to cover more than just chapter headings. However, it is not as straightforward as you might think. When alternation is used, the largest possible expression on either side of the '\|' character is matched. You might think that the following expression match either 'Chapter' or 'Section' followed by one or two digits occurring at the beginning and ending of a line:

Example

"^Chapter|Section [1-9][0-9]{0,1}$"

Unfortunately, the regular expression shown above matches either the word 'Chapter' at the beginning of a line, or 'Section' and whatever numbers follow that, at the end of the line. If the input string is 'Chapter 22', the expression shown above only matches the word 'Chapter'. If the input string is 'Section 22', the expression matches 'Section 22'. But that is not the intent here so there must be a way to make that regular expression more responsive to what you're trying to do and there is.

You can use parentheses to limit the scope of the alternation, that is, make sure that it applies only to the two words, 'Chapter' and 'Section'. However, parentheses are also used to create subexpressions and possibly capture them for later use, something that is covered in the section on backreferences. By taking the regular expressions shown above and adding parentheses in the appropriate places, you can make the regular expression match either 'Chapter 1' or 'Section 3'. 

The following regular expression uses parentheses to group 'Chapter' and 'Section' so the expression works properly.

"^(Chapter|Section) \[1-9]\[0-9]{0,1}$"

Although these expressions work properly, the parentheses around 'Chapter\|Section' also cause either of the two matching words to be captured for future use. Since there is only one set of parentheses in the expression shown above, there is only one captured submatch. This submatch can be referred to using the Submatches collection.

In the above example, you merely want to use the parentheses to group a choice between the words 'Chapter' and 'Section'. To prevent the match from being saved for possible later use, place '?:' before the regular expression pattern inside the parentheses. The following modification provides the same capability without saving the submatch:

"^(?:Chapter|Section) \[1-9]\[0-9]{0,1}$"

In addition to the '?:' metacharacters, there are two other non-capturing metacharacters used for something called lookahead matches. A positive lookahead, specified using ?=, matches the search string at any point where a matching regular expression pattern in parentheses begins. A negative lookahead, specified using '?!', matches the search string at any point where a string not matching the regular expression pattern begins.

For example, suppose you have a document containing references to Windows 3.1, Windows 95, Windows 98, and Windows NT. Suppose further that you need to update the document by finding all the references to Windows 95, Windows 98, and Windows NT and changing those reference to Windows 2000. You can use the following regular expression, which is an example of a positive lookahead, to match Windows 95, Windows 98, and Windows NT:

"Windows(?=95 |98 |NT )"

Once the match is found, the search for the next match begins immediately following the matched text, not including the characters included in the look-ahead. For example, if the expressions shown above matched 'Windows 98', the search resumes after 'Windows' not after '98'.

# <a name="Recipes"></a>Recipes

#### Finding similar words

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "My cat found a dead bat over the mat"
DIM cbsPattern AS CBSTR = $"\b[bcm]at\b"
'-or- DIM cbsPattern AS CBSTR = $"\b(b|c|m)at\b"
pRegExp.Execute(cbsText, cbsPattern)
FOR i AS LONG = 0 TO pRegExp.MatchesCount - 1
   print pRegExp.MatchValue(i)
NEXT
```

As an alternative, we can use the operator "\|":

```
DIM cbsPattern AS CBSTR = $"\b(b|c|m)at\b"
```

\\b is a word boundary, \[bcm] one of b, c, or m, followed by at and finally another word boundary (\\b),

```
Output:
cat
bat
mat
```

#### Finding variations on words

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "Hello, my name is John Doe, what's your name?"
DIM cbsPattern AS CBSTR = $"\bJoh?n(athan)? Doe\b"
pRegExp.Execute(cbsText, cbsPattern)
IF pRegExp.MatchesCount THEN print pRegExp.MatchValue
```

\\b is a word boundary (matches whole words only), Jo is the common part of John, Jon and Jonathan, h? is optional (it is followed by the metacharacter ?), n is a letter common in the three names, (anathan)? is an optional group of characters that will match Jonathan, a space followed by Doe, and another word boundary (\\b).

```
Output:
John Doe
```

#### Removing characters

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbs AS CBSTR
cbs = pRegExp.RemoveStr("abacadabra", "[abc]", TRUE) ' - prints "dr"
print cbs
```

Removes any of the characters, a, b or c, found in the string.

```
Output:
dr
```

#### Removing words

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbs AS CBSTR
cbs = pRegExp.RemoveStr("abacadabra", "ab")
print cbs
```
```
Output:
acadra
```

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbs AS CBSTR
cbs = pRegExp.RemoveStr("World, worldx, world", $"\bworld\b", TRUE)
print cbs
```

\\b is word boundary. TRUE in the third parameter indicates that it must ignore the case.

```
Output:
, worldx,
```

#### Replacing characters

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbs AS CBSTR
cbs = pRegExp.ReplaceStr("abacadabra", "[bac]", "*")
print cbs
```

Replaces any of the characters, a, b or c, found in the string with an asterisk.

```
Output:
*****d**r*
```

#### Replacing words

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbs AS CBSTR
cbs = pRegExp.ReplaceStr("Hello World", "World", "Earth")
print cbs
```

```
Output:
Hello, Earth
```

Replaces any occurrence of two consecutive identical words in a string of text with a single occurrence of the same word. The TRUE in the fourth parameter indicates that the operation is not case sensitive.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "Is is the cost of of gasoline going up up?"
DIM cbsPattern AS CBSTR = $"\b([a-z]+) \1\b"
DIM cbs AS CBSTR = pRegExp.ReplaceStr(cbsText, cbsPattern, "$1", TRUE)
print cbs
```

```
Output:
Is the cost of gasoline going up?
```

Adds an space after the dots that are immediately followed by a word.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a text.With some dots.Between words. This one not."
DIM cbsPattern AS CBSTR = "(\.)(\w)"
DIM cbs AS CBSTR = pRegExp.ReplaceStr(cbsText, cbsPattern, "$1 $2")
print cbs
```

```
Output:
This is a text. With some dots. Between words. This one not."
```

Changes the format of a telephone number.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
PRINT pRegExp.ReplaceStr("555-123-4567", "(\d{3})-(\d{3})-(\d{4})", "($1) $2-$3")
```

```
Output:
(555) 123-4567
```

Swaps to words and removes the comma.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
PRINT pRegExp.ReplaceStr("0000.34500044", $"\b0{1,}\.", ".")
```

```
Output:
Paul Squires
```

Removes leading zeroes.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
PRINT pRegExp.ReplaceStr("0000.34500044", $"\b0{1,}\.", ".")
```

```
Output:
.34500044
```

Replaces everything between quotes with three asterisks.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a ""quoted"" string"
DIm cbsPattern AS CBSTR = """[^""]*"""
PRINT pRegExp.ReplaceStr(cbsText, cbsPattern, "***")
```

```
Output:
This is a "***"" string
```

Replaces a tab with a pipe character.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a " & CHR(9) & " test string"
'DIM cbsText AS CBSTR = !"This is a \t test string" ' alternative way
DIm cbsPattern AS CBSTR = $"\t"
PRINT pRegExp.ReplaceStr(cbsText, cbsPattern, "|")
```

```
Output:
This is a | test string
```

#### Searching words and substrings

Searches for the first occurence of a word that stars with a letter ans is followed by three numbers. If found, it returns its position.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM nPos AS LONG = pRegExp.Find(cbsText, cbsPattern)
PRINT nPos
```

```
Output: 11
```

Searches all the occurences of a word that stars with a letter ans is followed by three numbers. Returns a list of comma separated "index, length" value pairs. The pairs are separated by a semicolon.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "blah blah a234 blah blah x345 blah blah"
DIM cbsPattern AS CBSTR = "[a-z][0-9][0-9][0-9]"
DIM cbsOut AS CBSTR
cbsOut = pRegExp.FindEx(cbsText, cbsPattern)
print cbsOut
```

```
Output: 11,4; 26,4
```

#### Extracting words and substrings

Searches for the first occurrence of a word. Case-sensitive and not global.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a test string"
DIM cbsPattern AS CBSTR = "test"
IF pRegExp.Execute(cbsText, cbsPattern, FALSE, FALSE) THEN PRINT pRegExp.MatchValue
```

```
Output:
test
```

Searches for the first occurrence of a word. Case-insensitive and not global.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a test string"
DIM cbsPattern AS CBSTR = "Test"
IF pRegExp.Execute(cbsText, cbsPattern, TRUE, FALSE) THEN PRINT pRegExp.MatchValue
```

```
Output:
test
```

Searches for the first occurrence of a substring. Case-sensitive and not global.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a test string"
DIM cbsPattern AS CBSTR = "is a test"
IF pRegExp.Execute(cbsText, cbsPattern, FALSE) THEN PRINT pRegExp.MatchValue
```

```
Output:
is a test
```

Searches for the first occurrence of a substring. Case-insensitive and not global.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a test string"
DIM cbsPattern AS CBSTR = "is a test"
IF pRegExp.Execute(cbsText, cbsPattern, TRUE, FALSE) THEN PRINT pRegExp.MatchValue
```

```
Output:
is a test
```

Searches for the all the occurrences of a word. Case-sensitive and global.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a test testx string"
DIM cbsPattern AS CBSTR = "test"
IF pRegExp.Execute(cbsText, cbsPattern, FALSE) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
```

```
Output:
test
test
```

Searches for the all the occurrences of a word. Case-insensitive and global.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a test testx string"
DIM cbsPattern AS CBSTR = "Test"
IF pRegExp.Execute(cbsText, cbsPattern, TRUE) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
```

```
Output:
test
test
```

Searches for the all the occurrences of a whole word. Case-insensitive and global.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "This is a test testx string"
DIM cbsPattern AS CBSTR = $"\bTest\b"
IF pRegExp.Execute(cbsText, cbsPattern, TRUE) THEN
   FOR i AS LONG = 0 TO pRegExp.MatchesCount
      PRINT pRegExp.MatchValue(i)
   NEXT
END IF
```

```
Output:
test
```

Case sensitive, double search (c.t and d.g), whole words. Retrieves cut, cat, i.e. whole words with three letters that begin with c and end with t or begin with d and end with g.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
#INCLUDE ONCE "Afx/CWSTR.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "I have a cat and a dog, because I love cats and dogs"
DIM cbsPattern AS CBSTR = "c.t|d.g"
pRegExp.Execute(cbsText, cbsPattern, FALSE)
FOR i AS LONG = 0 TO pRegExp.MatchesCount - 1
   PRINT pRegExp.MatchValue(i)
NEXT
```

```
Output:
cat
dog
cat
dog
```

Case insensitive, double search (c.t and d.g), whole words. Retrieves cut, cat, i.e. whole words with three letters that begin with c and end with t or begin with d and end with g.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
#include once "afx/cwstr.inc"
USING Afx

' // With this constructor we set the pattern, ignore case and global
DIM pRegExp AS CRegExp = CRegExp($"\bc.t\b|\bd.g\b", TRUE, TRUE)
pRegExp.Execute("I have a cat and a dog, because I love cats and dogs")
FOR i AS LONG = 0 TO pRegExp.MatchCount - 1
   PRINT pRegExp.MatchValue(i)
NEXT
```

```
Output:
cat
dog
cat
dog
```

We can search for more than a word at the same time.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
#INCLUDE ONCE "Afx/CWSTR.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "The contests will be in January, July and November"
DIm cbsPattern AS CBSTR = $"\b(january|february|march|april|may|june|july|" & _
    $"august|september|october|november|december)\b"
PRINT pRegExp.Execute(cbsText, cbsPattern, TRUE)
For i AS LONG = 0 TO pRegExp.MatchesCount - 1
   PRINT pRegExp.MatchValue(i)
NEXT
```

```
Output:
January
July
September
```

Extracts a quoted string.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "29:Sep:2017 ""This is an Example!"""
DIM cbsPattern AS CBSTR = """.*?"""
PRINT pRegExp.Execute(cbsText, cbsPattern)
PRINT pRegExp.MatchValue
```

```
Output:
"This is an Example!"
```

Extracts all the alphabetic words.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "29:Sep:2017 ""This is an Example!"""
DIM cbsPattern AS CBSTR = "(?:[a-z][a-z]+)"
pRegExp.Execute(cbsText, cbsPattern, TRUE)
FOR i AS LONG = 0 TO pRegExp.MatchesCount - 1
   PRINT pRegExp.MatchValue(i)
NEXT
```

```
Output:
Sep
This
is
an
Example
```

Extracts the year.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "29:Sep:2017 ""This is an Example!"""
DIM cbsPattern AS CBSTR = $"((?:(?:[1]{1}\d{1}\d{1}\d{1})|(?:[2]{1}\d{3})))(?![\d])"
pRegExp.Execute(cbsText, cbsPattern)
PRINT pRegExp.MatchValue
```

```
Output:
2017
```

Extracts integers.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "29:Sep:2017 ""This is an Example!"""
DIM cbsPattern AS CBSTR = $"(\d+)"
pRegExp.Execute(cbsText, cbsPattern)
FOR i AS LONG = 0 TO pRegExp.MatchesCount
   PRINT pRegExp.MatchValue(i)
NEXT
```

```
Output:
29
2017
```

Extracts text between parentheses.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp

' // extract the first match
DIM cbsText AS CBSTR = "blah blah (text beween parentheses) blah blah"
DIM cbs AS CBSTR = pRegExp.Extract(cbsText, "([^(]*?)(?=\))")
PRINT cbs

' // extract the first match after the 11th position
cbsText = "blah (xxx) blah (text beween parentheses) blah blah"
cbs = pRegExp.Extract(11, cbsText, "([^(]*?)(?=\))")
PRINT cbs
```

```
Output:
"text between parentheses"
```

Extracts text between curly braces.

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp

' // extract the first match
DIM cbsText AS CBSTR = "blah blah {text beween curly braces} blah blah"
DIM cbs AS CBSTR = pRegExp.Extract(cbsText, "([^{]*?)(?=\})")
PRINT cbs

' // extract the first match after the 11th position
cbsText = "blah (xxx) blah (text beween parentheses) blah blah"
cbs = pRegExp.Extract(11, cbsText, "([^{]*?)(?=\})")
PRINT cbs
```

```
Output:
"text between curly braces"
```

#### Checking if a string is numeric

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "1.2345678901234567e+029"
DIM cbsPattern AS CBSTR = "^[\+\-]?\d*\.?\d+(?:[Ee][\+\-]?\d+)?$"
PRINT pRegExp.Test(cbsText, cbsPattern)
```

```
Output:
True
```

```
Pattern: "^[\+\-]?\d*\.?\d+(?:[Ee][\+\-]?\d+)?$"
```

The initial "^" and the final "$" match the start and the end of the string, to ensure the check spans the whole string.

The "\[\\+\\-]?" part is the inital plus or minus sign with the "?" multiplier that allows zero or one instance of it.

The "\\d*" is a digit, zero or more times.

"\\\.?" is the decimal point, zero or one time.

The "\\d+" part matches a digit one or more times.

The "(?:\[Ee]\[\\+\\-]?\\d+)?" matches "e+", "e-", "E+" or "E-" followed by one or more digits, with the "?" multiplier that allows zero or one instance of it.

#### Checking if an url is valid

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "https://www.google.es/search?q=msdn+jscript+check+if+an+url+is+valid&client=firefox-b&dcr=0&ei=hM_OWfb2BMzSU-mhr7AG&start=20&sa=N&biw=947&bih=394"
DIM cbsPattern AS CBSTR = "(ftp|http|https):\/\/(\w+:{0,1}\w*@)?(\S+)(:[0-9]+)?(\/|\/([\w#!:.?+=&%@!\-\/]))?"
PRINT pRegExp.Test(cbsText, cbsPattern)
```

#### Breaking down an URI in its parts

```
#INCLUDE ONCE "Afx/CRegExp.inc"
USING Afx

' // Breaks down a URI down to the protocol (ftp, http, and so on), the domain
' // address, and the page/path
DIM pRegExp AS CRegExp
DIM cbsText AS CBSTR = "http://msdn.microsoft.com:80/scripting/default.htm"
DIM cbsPattern AS CBSTR = $"(\w+):\/\/([^/:]+)(:\d*)?([^# ]*)"
pRegExp.Execute(cbsText, cbsPattern)
FOR i AS LONG = 0 TO pRegExp.SubMatchesCount - 1
   PRINT pRegExp.SubMatchValue(0, i)
NEXT
```

```
Output:
http
msdn.microsoft.com
:80
/scripting/default.htm
```
