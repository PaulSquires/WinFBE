' ########################################################################################
' Microsoft Windows
' Contents: ADO - Stream example
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates the use of the Stream object.
' Note: Error checking ommited for brevity.
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'#CONSOLE ON
#define UNICODE
#include "Afx/CADODB/CADODB.inc"
USING Afx

' // Open a stream in memory
DIM pStream AS CAdoStream
pStream.Type_ = adTypeText
pStream.LineSeparator = adCRLF
pStream.Open

' // Write some text to it
pStream.WriteText "This is a test string", adWriteLine
pStream.WriteText "This is another test string", adWriteLine

' // Set the position at the beginning of the file
pStream.Position = 0
' // Read the lines
DIM cbsText AS CBSTR = pStream.ReadText(adReadLine)
print cbsText
cbsText = pStream.ReadText(adReadLine)
print cbsText

' // Save the contents to a file
pStream.SaveToFile "TestStream.txt", adSaveCreateOverWrite
pStream.Close

PRINT
PRINT "Press any key..."
SLEEP
