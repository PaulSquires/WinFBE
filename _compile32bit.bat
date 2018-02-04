set FBLIBDIR=win32
set FBCPATH=x:\FB\FreeBASIC-1.06.0-win32

cd src
%FBCPATH%\fbc WinFBE.bas -x ..\WinFBE32.exe WinFBE.rc -s gui
cd ..
 
@if ErrorLevel 1 goto compileerror
@echo SUCCESS!! WinFBE (32 bit) built.

goto exit

:compileerror
@echo *** COMPILE ERROR ***

:exit
@pause

