set FBLIBDIR=win32
rem set FBCPATH=x:\PSS121\FB\WinFBE_Suite\FreeBASIC-1.07.1-gcc-5.2
set FBCPATH=%~d0\PSS121\FB\WinFBE_Suite\FreeBASIC-1.07.1-gcc-5.2

cd src
%FBCPATH%\fbc32 WinFBE.bas -x ..\WinFBE32.exe WinFBE.rc -s gui
cd ..
rem X:\PSS121\FB\WinFBE_Suite-Editor\Tools\upx-3.95-win32\upx.exe WinFBE32.exe
 
