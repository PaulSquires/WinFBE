' ########################################################################################
' Microsoft Windows
' File: MatrixEquals.bas
' Contents: GDI+ - MatrixEquals example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'#CONSOLE ON
#define UNICODE
#INCLUDE ONCE "Afx/CGdiPlus/CGdiPlus.inc"
USING Afx

' // Must be constructed with NEW to we able to delete them before the call to AfxGdipShutdown
DIM pMat1 AS CGpMatrix PTR = NEW CGpMatrix(1.0, 2.0, 1.0, 1.0, 2.0, 2.0)
DIM pMat2 AS CGpMatrix PTR = NEW CGpMatrix(1.0, 2.0, 1.0, 1.0, 2.0, 2.0)

IF pMat1->Equals(pMat2) THEN
   PRINT "They are the same"
ELSE
   PRINT "They are the different"
END IF

' // Must be deleted before the call to AfxGdipShutdown
Delete pMat1
Delete pMat2

PRINT
PRINT "Press any key"
SLEEP
