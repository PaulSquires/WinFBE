set FBLIBDIR=win64
set FBCPATH=X:\PSS121\FB\WinFBE_Suite\FreeBASIC-1.07.1-gcc-5.2

cd src
%FBCPATH%\fbc64 WinFBE.bas -x ..\WinFBE64.exe WinFBE.rc -s gui
cd ..
X:\PSS121\FB\Tools\upx-3.95-win32\upx.exe WinFBE64.exe


