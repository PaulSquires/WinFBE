'    LangEdit - Language Editor for the WinFBE Editor
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
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



type clsLanguage
   private:
   
   public:
      phrases(MAXPHRASES) as CWSTR
      LanguageName as CWSTR   ' english, french, german, etc
      FullFileName as CWSTR
      declare function LoadLanguage() as Boolean
end type


function clsLanguage.LoadLanguage() as Boolean

   ' default that the file failed to load
   Function = False

   Dim as CBSTR wst, wKey, wData
   Dim nKey  As Long
   Dim nData As Long  
   Dim i     As Long
    
   dim pStream AS CTextStream
   if pStream.OpenUnicode(this.FullFileName) <> S_OK then exit function

   do until pStream.EOS
      wst = pStream.ReadLine
      
      If Len(wst) = 0 Then Continue Do
      If Left(wst, 1) = "'" Then Continue Do
      
      i = Instr(wst, ":")
      If i = 0 Then Continue Do
      
      wKey = "": wData = "": nData = 0

      wKey  = Left(wst, i-1)
      wData = Mid(**wst, i+1)    ' MID causes problems with Chinese data so ** is used.
      
      nKey  = Val(wKey)
      nData = Val(wData)

      If Ucase(wKey) = "MAXIMUM" Then
         ' Our class array is fixed at the #Define MAXPHRASES. That value should be
         ' at least equal to or greater than the MAXIMUM value as defined in the file.
         continue do
      end if

      ' this should be a key/value pair line in the format:
      ' 00001:value
      ' Ensure that we add the value to the array within the valid
      ' boundaries of the array.
      If (nKey >= LBound(this.phrases)) AndAlso (nKey <= Ubound(this.phrases)) Then
         ' Remove any comments from end of the line. Comments begin with
         ' a semicolon character.
         this.phrases(nKey) = rtrim(AfxStrParse(wData, 1, ";"), any chr(9,32))
      End If
         
   Loop
   pStream.Close
   
   function = true   ' successful loading the language file
end function


dim shared gLangEnglish as clsLanguage ptr
