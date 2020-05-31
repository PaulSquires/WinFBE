set FBLIBDIR=win64
rem set FBCPATH=X:\PSS121\FB\WinFBE_Suite\FreeBASIC-1.07.1-gcc-8.4
set FBCPATH=%~d0\PSS121\FB\WinFBE_Suite\FreeBASIC-1.07.1-gcc-8.4

cd src
%FBCPATH%\fbc64 WinFBE.bas -x ..\WinFBE64.exe WinFBE.rc -s gui
cd ..
rem X:\PSS121\FB\WinFBE_Suite-Editor\Tools\upx-3.95-win32\upx.exe WinFBE64.exe


