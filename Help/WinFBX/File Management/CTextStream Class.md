# CTextStream Class

Allows to read and write sequential text files (sometimes referred to as a text stream).
Works with ASCII and Unicode.
Works with Windows CRLF files and with Linux LF files.

**Include file**: CTextStream.inc

### Constructors

Besides the default constructor, that initializes the COM library and creates an instance of the ITextStream interface, this additional constructor allows to create an instance of the class by passing a pointer to an existing ITextStream interface that becomes attached to the class (therefore, don't release the passed pointer, it will be released by the CTextStream class).

```
CONSTRUCTOR CTextStream (BYVAL pstm AS Afx_ITextStream PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstm* | Pointer to an instance of the Afx_ITextStream interface. |

This allows to work with the standard StdIn, StdErr and StdOut streams using the methods of the **CTextStream** class.

```
#include "Afx/AfxFile.inc"
' // Create an instance of the CTextStream class initialized with a pointer to the standard StdOut stream
DIM pTxtStm AS CTextStream = CFileSys().GetStdOutStream(TRUE)   ' or FALSE to work with ansi
' // Write a string and an end of line to the stream
pTxtStm.WriteLine "This is a test."
```

### Operators

| Name       | Description |
| ---------- | ----------- |
| LET | Initializes the class from an existing stream and attaches it. |
| CAST | Returns a pointer to the underlying ITextStream interface of the text stream object. |

```
OPERATOR LET (BYVAL pstm AS Afx_ITextStream PTR)
OPERATOR CAST () AS Afx_ITextStream PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstm* | A pointer to the ITextStream interface of an existing text stream that will be attached to the class. |

### Methods and properties

| Name       | Description |
| ---------- | ----------- |
| [Create](#Create) | Creates a file and returns a TextStream object that can be used to read from or write to the file. |
| [Open](#Open) | Opens a file and returns a TextStream object that can be used to read from, write to, or append to the file. |
| [OpenUnicode](#OpenUnicode) | Opens a file and returns a TextStream object that can be used to read from, write to, or append to the file. |
| [OpenForInputA](#OpenForInputA) | Opens a file for input and returns a TextStream object that can be used to read from the file. |
| [OpenForInputW](#OpenForInputW) | Opens a file for input and returns a TextStream object that can be used to read from the file. |
| [OpenForOutputA](#OpenForOutputA) | Opens a file for putput and returns a TextStream object that can be used to write to the file. |
| [OpenForOutputW](#OpenForOutputW) | Opens a file for putput and returns a TextStream object that can be used to write to the file. |
| [OpenForAppendA](#OpenForAppendA) | Opens a file for append and returns a TextStream object that can be used to write to the file. |
| [OpenForAppendW](#OpenForAppendW) | Opens a file for append and returns a TextStream object that can be used to write to the file. |
| [Close](#Close) | Closes an open TextStream file. |
| [EOL](#EOL) | Returns TRUE if the file pointer is positioned immediately before the end-of-line marker in a TextStream file; FALSE if it is not. |
| [EOS](#EOS) | Returns TRUE if the file pointer is at the end of a TextStream file; FALSE if it is not. |
| [Read](#Read) | Reads a specified number of characters from a TextStream file and returns the resulting string. |
| [ReadLine](#ReadLine) | Reads an entire line (up to, but not including, the newline character) from a TextStream file and returns the resulting string. |
| [ReadAll](#ReadAll) | Reads an entire TextStream file and returns the resulting string. |
| [Write](#Write) | Writes a specified string to a TextStream file. |
| [WriteLine](#WriteLine) | Writes a specified string and newline character to a TextStream file. |
| [WriteBlankLines](#WriteBlankLines) | Writes a specified number of newline characters to a TextStream file. |
| [Skip](#Skip) | Skips a specified number of characters when reading a TextStream file. |
| [SkipLine](#SkipLine) | Skips the next line when reading a TextStream file. |
| [Line](#Line) | Returns the current line number in a TextStream file. |
| [Column](#Column) | Returns the column number of the current character position in a TextStream file. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |


# <a name="Create"></a>Create

Creates a specified file name and returns a TextStream object that can be used to read from or write to the file.

```
FUNCTION Create (BYREF cbsFileName AS CBSTR, _
   BYVAL bOverwrite AS BOOLEAN = TRUE, _
   BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFileName* | CBSTR. String expression that identifies the file to create. |
| *bOverwrite* | Boolean value that indicates whether you can overwrite an existing file. The value is true if the file can be overwritten, false if it can't be overwritten. If omitted, existing files are not overwritten. |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it's created as an ASCII file. If omitted, an ASCII file is assumed. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Create a text stream
DIM cbsFile AS CBSTR = ExePath & "\Test.txt"
pTxtStm.Create(cbsFile, TRUE)
' // Write a string and an end of line to the stream
pTxtStm.WriteLine "This is a test."
' // Write more strings
pTxtStm.WriteLine "This is a string."
pTxtStm.WriteLine "This is a second string."
pTxtStm.WriteLine "This is the end line."

PRINT
PRINT "Press any key o end..."
SLEEP
```
# <a name="Open"></a>Open

Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.

```
FUNCTION Open (BYREF cbsFileName AS CBSTR, _
   BYVAL IOMode AS LONG = 1, _
   BYVAL bCreate AS BOOLEAN = FALSE, _
   BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFileName* | CBSTR. String expression that identifies the file to open. |
| *IOMode* | LONG. Can be one of three constants: IOMode_ForReading, IOMode_ForWriting, or IOMode_ForAppending. |
| *bCreate* | Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created. |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it's created as an ASCII file. If omitted, an ASCII file is assumed. |

The *IOMode* argument can have any of the following settings:

| Constant | Value | Description |
| -------- | ----- | ----------- |
| IOMode_ForReading   | 1 | Open a file for reading only. You can't write to this file. |
| IOMode_ForWriting   | 2 | Open a file for writing. |
| IOMode_ForAppending | 8 | Open a file and write to the end of the file. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
'#CONSOLE ON
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Open file as a text stream
DIM cbsFile AS CBSTR = ExePath & "\Test.txt"
pTxtStm.Open(cbsFile, IOMode_ForReading)
' // Read the file sequentially
DIM s AS STRING
DO UNTIL pTxtStm.EOS
   s = pTxtStm.ReadLine
   PRINT s
LOOP
pTxtStm.Close

PRINT "Press any key..."
SLEEP
```

# <a name="OpenUnicode"></a>OpenUnicode

Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.

```
FUNCTION OpenUnicode (BYREF cbsFileName AS CBSTR, _
   BYVAL IOMode AS LONG = 1, _
   BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFileName* | CBSTR. String expression that identifies the file to open. |
| *IOMode* | LONG. Can be one of three constants: IOMode_ForReading, IOMode_ForWriting, or IOMode_ForAppending. |
| *bCreate* | Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created. |

The *IOMode* argument can have any of the following settings:

| Constant | Value | Description |
| -------- | ----- | ----------- |
| IOMode_ForReading   | 1 | Open a file for reading only. You can't write to this file. |
| IOMode_ForWriting   | 2 | Open a file for writing. |
| IOMode_ForAppending | 8 | Open a file and write to the end of the file. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
'#CONSOLE ON
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Open file as a text stream
DIM cbsFile AS CBSTR = ExePath & "\TestW.txt"
IF pTxtStm.OpenUnicode(cbsFile, IOMode_ForReading) = S_OK THEN
   ' // Read the file sequentially
   DIM cbs AS CBSTR
   DO UNTIL pTxtStm.EOS
      cbs = pTxtStm.ReadLine
      PRINT cbs
   LOOP
   pTxtStm.Close
END IF

PRINT "Press any key..."
SLEEP
```

# <a name="OpenForInputA"></a>OpenForInputA

Opens a file for input and returns a TextStream object that can be used to read from the file. (Ansi version.)

```
FUNCTION OpenForInputA (BYREF cbsFileName AS CBSTR, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFileName* | CBSTR. String expression that identifies the file to open. |
| *bCreate* | Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
'#CONSOLE ON
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // ' // Open file for input as a text stream
IF pTxtStm.OpenForInputA(ExePath & "\Test.txt") = S_OK THEN
   ' // Read the file sequentially
   DIM s AS STRING
   DO UNTIL pTxtStm.EOS
      s = pTxtStm.ReadLine
      PRINT s
   LOOP
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="OpenForInputW"></a>OpenForInputW

Opens a file for input and returns a TextStream object that can be used to read from the file. (Unicode version.)

```
FUNCTION OpenForInputW (BYREF cbsFileName AS CBSTR, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFileName* | CBSTR. String expression that identifies the file to open. |
| *bCreate* | Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
'#CONSOLE ON
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // ' // Open file for input as a text stream
IF pTxtStm.OpenForInputW(ExePath & "\TestW.txt") = S_OK THEN
   ' // Read the file sequentially
   DIM cbs AS CBSTR
   DO UNTIL pTxtStm.EOS
      cbs = pTxtStm.ReadLine
      PRINT cbs
   LOOP
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="OpenForOutputA"></a>OpenForOutputA

Opens a file for output and returns a TextStream object that can be used to write to the file. (Ansi version.)

```
FUNCTION OpenForOutputA (BYREF cbsFileName AS CBSTR, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFileName* | CBSTR. String expression that identifies the file to open. |
| *bCreate* | Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
'#CONSOLE ON
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Open file for output as a text stream
IF pTxtStm.OpenForOutputA(ExePath & "\Test.txt") = S_OK THEN
   ' // Write a string and an end of line to the stream
   pTxtStm.WriteLine "This is a test."
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="OpenForOutputW"></a>OpenForOutputW

Opens a file for output and returns a TextStream object that can be used to write to the file. (Unicode version.)

```
FUNCTION OpenForOutputW (BYREF cbsFileName AS CBSTR, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFileName* | CBSTR. String expression that identifies the file to open. |
| *bCreate* | Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
'#CONSOLE ON
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Open file for output as a text stream
IF pTxtStm.OpenForOutputW(ExePath & "\TestW.txt") = S_OK THEN
   ' // Write a string and an end of line to the stream
   pTxtStm.WriteLine "This is a test."
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```


# <a name="OpenForAppendA"></a>OpenForAppendA

Opens a file for append and returns a TextStream object that can be used to write to the file. (Ansi version.)

```
FUNCTION OpenForOutputA (BYREF cbsFileName AS CBSTR, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFileName* | CBSTR. String expression that identifies the file to open. |
| *bCreate* | Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
'#CONSOLE ON
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Open file for append as a text stream
IF pTxtStm.OpenForAppendA(ExePath & "\Test.txt") = S_OK THEN
   ' // Write a string and an end of line to the stream
   pTxtStm.WriteLine "This is a test."
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="OpenForAppendW"></a>OpenForAppendW

Opens a file for append and returns a TextStream object that can be used to write to the file. (Unicode version.)

```
FUNCTION OpenForAppendW (BYREF cbsFileName AS CBSTR, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsFileName* | CBSTR. String expression that identifies the file to open. |
| *bCreate* | Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created, False if it isn't created. If omitted, a new file isn't created. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

```
'#CONSOLE ON
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Open file for append as a text stream
IF pTxtStm.OpenForAppendW(ExePath & "\TestW.txt") = S_OK THEN
   ' // Write a string and an end of line to the stream
   pTxtStm.WriteLine "This is a test."
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="Close"></a>Close

Closes an open TextStream file.

```
FUNCTION Close () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="EOL"></a>EOL

Returns TRUE if the file pointer is positioned immediately before the end-of-line marker in a TextStream file; FALSE if it is not.

```
FUNCTION EOL () AS BOOLEAN
```

#### Remark

The EOL (End of Line) property applies only to TextStream files that are open for reading; otherwise, an error occurs.

# <a name="EOS"></a>EOS

Returns TRUE if the file pointer is at the end of a TextStream file; FALSE if it is not.

```
FUNCTION EOS () AS BOOLEAN
```

#### Remark

The EOS (End os Stream) property applies only to TextStream files that are open for reading, otherwise, an error occurs.


# <a name="Read"></a>Read

Reads a specified number of characters from a TextStream file and returns the resulting string.

```
FUNCTION Read (BYVAL numChars AS LONG) AS CBSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *numChars* | LONG. Number of characters you want to read from the file. |

#### Return value

CBSTR. A string with the characters read.

# <a name="ReadLine"></a>ReadLine

Reads an entire line (up to, but not including, the newline character) from a TextStream file and returns the resulting string.

```
FUNCTION ReadLine  () AS CBSTR
```

#### Return value

CBSTR. A string with the content of the line read.

#### Example

```
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Open file as a text stream
DIM cbsFile AS CBSTR = ExePath & "\Test.txt"
IF pTxtStm.Open(cbsFile, IOMode_ForReading) = S_OK THEN
   ' // Read the file sequentially
   DO
     IF pTxtStm.EOS THEN EXIT DO
     DIM curLine AS LONG = pTxtStm.Line
     DIM cbsText AS CBSTR = pTxtStm.ReadLine
     PRINT "Line " & STR(curLine) & ": " & cbsText
   LOOP
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="ReadAll"></a>ReadAll

Reads an entire TextStream file and returns the resulting string.

```
FUNCTION ReadAll  () AS CBSTR
```

#### Return value

CBSTR. A string with the content of all the file.

#### Example

```
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Open file as a text stream
DIM cbsFile AS CBSTR = ExePath & "\Test.txt"
IF pTxtStm.Open(cbsFile, IOMode_ForReading) = S_OK THEN
   ' // Read all the contents of the file
   DIM cbsText AS CBSTR = pTxtStm.ReadAll
   PRINT cbsText
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="Write"></a>Write

Writes a specified string to a TextStream file.

```
FUNCTION Write (BYREF cbsText AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsText* | CBSTR. The text you want to write to the file. |

#### Return value

CBSTR. The text you want to write to the file.

#### Example

```
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Create a text stream
DIM cbsFile AS CBSTR = ExePath & "\Test.txt"
IF pTxtStm.Create(cbsFile, TRUE) = S_OK THEN
   ' // Write strings
   pTxtStm.Write "This is a string."
   pTxtStm.Write "This is a second string."
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="WriteLine"></a>WriteLine

Writes a specified string and newline character to a TextStream file.

```
FUNCTION WriteLine (BYREF cbsText AS CBSTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsText* | CBSTR. The text you want to write to the file. If omitted, a newline character is written to the file. |

#### Return value

CBSTR. The text you want to write to the file.

#### Example

```
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Create a text stream
DIM cbsFile AS CBSTR = ExePath & "\Test.txt"
IF pTxtStm.Create(cbsFile, TRUE) = S_OK THEN
   ' // Write strings
   pTxtStm.WriteLine "This is a string."
   pTxtStm.WriteLine "This is a second string."
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="WriteBlankLines"></a>WriteBlankLines

Writes a specified string and newline character to a TextStream file.

```
FUNCTION WriteBlankLines (BYVAL numLines AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *numLines* | LONG. Number of newline characters you want to write to the file. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Create a text stream
DIM cbsFile AS CBSTR = ExePath & "\Test.txt"
IF pTxtStm.Create(cbsFile, TRUE) = S_OK THEN
   ' // Write a string and an end of line to the stream
   pTxtStm.WriteLine "This is a test."
   ' // Write more strings
   pTxtStm.Write "This is a string."
   pTxtStm.Write "This is a second string."
   ' // Write two blank lines (the first will serve as an end of line for the previous write instructions)
   pTxtStm.WriteBlankLines 2
   pTxtStm.WriteLine "This is the end line."
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="Skip"></a>Skip

Skips a specified number of characters when reading a TextStream file.

```
FUNCTION Skip (BYVAL numChars AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *numChars* | LONG. Number of characters  to skip when reading a file. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Example

```
#include "Afx/CTextStream.inc"
using Afx

' // Create an instance of the CTextStream class
DIM pTxtStm AS CTextStream
' // Open file as a text stream
DIM cbsFile AS CBSTR = ExePath & "\Test.txt"
IF pTxtStm.Open(cbsFile, IOMode_ForReading) = S_OK THEN
   ' // Read the file sequentially
   DO
     ' // Ext if we are at the end of the stream
     IF pTxtStm.EOS THEN EXIT DO
     ' // Current line
     DIM curLine AS LONG = pTxtStm.Line
     ' // Skip the 3rd line
     IF curLine = 3 THEN
        pTxtStm.SkipLine
        curLine + = 1
     END IF
     ' // Skip 10 characters
     pTxtStm.Skip 10
     ' // Current column
     DIM curColumn AS LONG = pTxtStm.Column
     ' // Read 5 characters
     DIM cbsText AS CBSTR = pTxtStm.Read(5)
     ' // Skip the rest of the line
     pTxtStm.SkipLine
     PRINT "Line " + WSTR(curLine) + ", Column " +  WSTR(curColumn) & ": " + cbsText
   LOOP
   '// Close the file
   pTxtStm.Close
END IF

PRINT "Press any key to end..."
SLEEP
```

# <a name="SkipLine"></a>SkipLine

Skips the next line when reading a TextStream file.

```
FUNCTION SkipLine () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="Line"></a>Line

Read-only property that returns the current line number in a TextStream file.

```
PROPERTY Line  () AS LONG
```

#### Return value

LONG. The current line number.

### Remarks

After a file is initially opened and before anything is written, Line is equal to 1.

# <a name="Column"></a>Column

Read-only property that returns the column number of the current character position in a TextStream file.

```
PROPERTY Column  () AS LONG
```

#### Return value

LONG. The current column number.

#### Remarks

After a newline character has been written, but before any other character is written, Column is equal to 1.

# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.
