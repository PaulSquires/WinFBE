set FBLIBDIR=win64
set FBCPATH=x:\FB\FreeBASIC-1.05.0-win64

%FBCPATH%\fbc WinFBE.bas WinFBE.rc -s gui

@if ErrorLevel 1 goto compileerror
@echo SUCCESS!! WinFBE (64 bit) built.

goto exit

:compileerror
@echo *** COMPILE ERROR ***

:exit
@pause

