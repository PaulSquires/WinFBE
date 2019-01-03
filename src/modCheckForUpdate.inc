'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2019 Paul Squires, PlanetSquires Software
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.


''
''  DoCheckForUpdates
''
''  Contact the PlanetSquires server and download the text file containing
''  the latest WinFBE version number. If an update exists then ask the user
''  if they wish to navigate to the Releases page on GitHub.
''
function DoCheckForUpdates() as long
   
   dim as CWSTR wszServerFile = "winfbe_version.txt"
   dim as CWSTR wszFilename
   dim as CWSTR wszLatestVersion
   dim as CWSTR wszMsg
   dim as string st
    
   DIM pWHttp AS CWinHttpRequest

   ' Open an HTTP connection to an HTTP resource
   pWHttp.Open "GET", "https://www.planetsquires.com/" & wszServerFile

   ' Send HTTP request and wait 5 seconds for response 
   pWHttp.Send
    
   DIM iSucceeded AS LONG = pWHttp.WaitForResponse(5)
   if iSucceeded then
      st = pWHttp.GetResponseBody
       
      ' Open a file stream and save the downloaded file
      DIM pFileStream AS CFileStream
      IF pFileStream.Open("winfbe_version.txt", STGM_CREATE OR STGM_WRITE) = S_OK then
         pFileStream.WriteTextA(st)
      END IF
   end if
    
   ' Open the downloaded file and check for the version number
   wszFilename = AfxGetExePathName + wszServerFile
   if AfxFileExists( wszFilename ) then
       
      dim as long f = freefile
      open wszFilename for input as #f
       
      do until eof(f)
         line input #f, st
         if left( st, 15 ) = "latest_version=" then
            wszLatestVersion = mid( st, 16 )
            exit do
         end if
      loop
          
      close #f
      AfxDeleteFile( wszFilename ) 
   else
      MessageBox( HWND_FRMMAIN, _
                  L(92,"Failed to retrieve update information"), _
                  L(94,"Software Update"), _
                  MB_ICONINFORMATION Or MB_OK )
      exit function
   end if

   if len(wszLatestVersion) then
      dim as long latestVersion = ConvertWinFBEversion(wszLatestVersion)
      dim as long installedVersion = ConvertWinFBEversion(APPVERSION)

      if installedVersion = latestVersion then
         wszMsg = L(96,"You are up to date!") & vbcrlf & _
                  "WinFBE v" & APPVERSION & " " & L(97,"is currently the newest version available.")
         MessageBox( HWND_FRMMAIN, wszMsg,L(94,"Software Update"), MB_ICONINFORMATION Or MB_OK )
         exit function
      
      elseif installedVersion < latestVersion then
         wszMsg = APPNAME & vbcrlf & _
                  L(98,"A new version is available.") & vbcrlf & vbcrlf & _
                  L(99,"Current") & ": " & APPVERSION & vbcrlf & _
                  L(107,"Available") & ": " & wszLatestVersion & vbcrlf & vbcrlf & _
                  L(137,"Do you wish to visit the download website?")
         If MessageBox( HWND_FRMMAIN, wszMsg, L(94,"Software Update"), MB_ICONQUESTION Or MB_YESNOCANCEL ) = IDYES Then 
            ShellExecute( NULL, "open", "https://github.com/PaulSquires/WinFBE/releases", Null, Null, SW_SHOWNORMAL )
         end if
         exit function
      end if
   end if
   
   function = 0
end function
