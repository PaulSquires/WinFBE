' ########################################################################################
' Microsoft Windows
' File: AfxMisc.inc
' Contents: Miscellaneous wrapper functions
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "Afx/CWmiCli.inc"
USING Afx

' ========================================================================================
' Retrieves the baseboard (also known as a motherboard or system board) serial number.
' ========================================================================================
PRIVATE FUNCTION AfxGetBaseBoardSerialNumber () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT SerialNumber FROM Win32_BaseBoard")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("SerialNumber")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the Bios serial number.
' ========================================================================================
PRIVATE FUNCTION AfxGetBiosSerialNumber () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT SerialNumber FROM Win32_BIOS")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("SerialNumber")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the manufacturer serial number.
' Contrarily to the serial number returned by AfxGetDiskDriveSerialNumber, this one won't
' change even if you format your hard drive.
' Requires Windows Vista+.
' ========================================================================================
PRIVATE FUNCTION AfxGetManufacturerSerialNumber () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT SerialNumber FROM Win32_PhysicalMedia")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("SerialNumber")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the disk drive serial number.
' Contrarily to the serial number returned by AfxGetManufacturerSerialNumber, that does not
' change, this one will change every time the hard drive is formatted.
' Requires Windows Vista+.
' ========================================================================================
PRIVATE FUNCTION AfxGetDiskDriveSerialNumber () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT SerialNumber FROM Win32_DiskDrive")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("SerialNumber")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the address width of the processor. On a 32-bit operating system, the value is
' 32 and on a 64-bit operating system it is 64. This PRIVATE FUNCTION can be used to determine if
' the processor is 32 or 64 bit.
' ========================================================================================
PRIVATE FUNCTION AfxGetAddressWidth () AS USHORT
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT AddressWidth FROM Win32_Processor")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("AddressWidth", @vRes) = S_OK THEN RETURN vRes.uiVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the system running on the Windows-based computer. The following list identifiers
' the returned value: "X86-based PC", "MIPS-based PC", "Alpha-based PC", "Power PC",
' "SH-x PC", "StrongARM PC", "64-bit Intel PC", "64-bit Alpha PC", "Unknown", "X86-Nec98 PC".
' ========================================================================================
PRIVATE FUNCTION AfxGetSystemType () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT SystemType FROM Win32_ComputerSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("SystemType")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the architecture of the operating system, as opposed to the processor.
' Example: "64 bits".
' Not available in Windows Server 2003, Windows 2000, Windows NT 4.0, Windows XP, and Windows Me/98/95.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSArchitecture () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT OSArchitecture FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("OSArchitecture")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the type of the computer in use, such as laptop, desktop, or Tablet.
' Not available in Windows Server 2003, Windows XP, Windows 2000, Windows NT 4.0, and Windows Me/98/95.
' Value   Meaning
' ------- --------------------------------------------
' 0 (&H0) Unspecified
' 1 (&H1) Desktop
' 2 (&H2) Mobile
' 3 (&H3) Workstation
' 4 (&H4) Enterprise Server
' 5 (&H5) Small Office and Home Office (SOHO) Server
' 6 (&H6) Appliance PC
' 7 (&H7) Performance Server
' 8 (&H8) Maximum
' ========================================================================================
PRIVATE FUNCTION AfxGetPCSystemType () AS USHORT
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT PCSystemType FROM Win32_ComputerSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("PCSystemType", @vRes) = S_OK THEN RETURN vRes.uiVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the name of the disk drive from which the Windows operating system starts.
' Example: "\Device\Harddisk0"
' ========================================================================================
PRIVATE FUNCTION AfxGetBootDevice () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT BootDevice FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("BootDevice")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Build number of an operating system. It can be used for more precise version information
' than product release version numbers. Example: "7601".
' ========================================================================================
PRIVATE FUNCTION AfxGetOSBuildNumber () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT BuildNumber FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("BuildNumber")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Type of build used for an operating system.
' Examples: "retail build", "checked build", "Multiprocessor Free".
' ========================================================================================
PRIVATE FUNCTION AfxGetOSBuildType () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT BuildType FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("BuildType")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Short description of the object—a one-line string. The string includes the operating
' system version. For example, "Microsoft Windows 7 Enterprise ". This property can be localized.
' Windows Vista and Windows 7: This property may contain trailing characters. For example,
' the string "Microsoft Windows 7 Enterprise " (trailing space included) may be necessary
' to retrieve information using this property.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSCaption () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("Caption")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Code page value an operating system uses. A code page contains a character table that an
' operating system uses to translate strings for different languages. The American National
' Standards Institute (ANSI) lists values that represent defined code pages. If an operating
' system does not use an ANSI code page, this member is set to 0 (zero). The CodeSet string
' can use a maximum of six characters to define the code page value.
' Example: "1255"
' ========================================================================================
PRIVATE FUNCTION AfxGetOSCodeSet () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT CodeSet FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("CodeSet")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Code for the country/region that an operating system uses. Values are based on international
' phone dialing prefixes—also referred to as IBM country/region codes. This property can
' use a maximum of six characters to define the country/region code value.
' Example: "1" (United States)
' ========================================================================================
PRIVATE FUNCTION AfxGetOSCountryCode () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT CountryCode FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("CountryCode")
END FUNCTION
' ========================================================================================

' ========================================================================================
' String that indicates the latest service pack installed on a computer. If no service pack
' is installed, the string is NULL. Example: "Service Pack 3".
' ========================================================================================
PRIVATE FUNCTION AfxGetOSCSDVersion () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT CSDVersion FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("CSDVersion")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Name of the scoping computer system.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSCSName () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT CSName FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("CSName")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Number, in minutes, an operating system is offset from Greenwich mean time (GMT).
' The number is positive, negative, or zero.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSCurrentTimeZone () AS USHORT
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT CurrentTimeZone FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("CurrentTimeZone", @vRes) = S_OK THEN RETURN vRes.uiVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Language identifier used by the operating system. A language identifier is a standard
' international numeric abbreviation for a country/region. Each language has a unique
' language identifier (LANGID), a 16-bit value that consists of a primary language
' identifier and a secondary language identifier.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSLocale () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT Locale FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("Locale")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Name of the registered user of the operating system. Example: "Ben Smith".
' ========================================================================================
PRIVATE FUNCTION AfxGetOSRegisteredUser () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT RegisteredUser FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("RegisteredUser")
END FUNCTION
' ========================================================================================

' ========================================================================================
' System directory of the operating system. Example: "C:\Windows\system32".
' ========================================================================================
PRIVATE FUNCTION AfxGetOSSystemDirectory () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT SystemDirectory FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("SystemDirectory")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Letter of the disk drive on which the operating system resides. Example: "C:".
' ========================================================================================
PRIVATE FUNCTION AfxGetOSSystemDrive () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT SystemDrive FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("SystemDrive")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Version number of the operating system. Example: "6.1.7601".
' ========================================================================================
PRIVATE FUNCTION AfxGetOSVersion () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT Version FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("Version")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Company name for the registered user of the operating system.
' Example: "Microsoft Corporation"
' ========================================================================================
PRIVATE FUNCTION AfxGetOSOrganization () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT Organization FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("Organization")
END FUNCTION
' ========================================================================================

' ========================================================================================
' The date in which the OS was installed.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSInstallDate () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT InstallDate FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("InstallDate")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Date and time the operating system was last restarted.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSLastBootUpTime () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("LastBootUpTime")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Operating system version of the local date and time-of-day.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSLocalDateTime () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT LocalDateTime FROM Win32_OperatingSystem")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("LocalDateTime")
END FUNCTION
' ========================================================================================

' ========================================================================================
' Encryption level for secure transactions: 40-bit, 128-bit, or n-bit.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSEncryptionLevel () AS ULONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT EncryptionLevel FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("EncryptionLevel", @vRes) = S_OK THEN RETURN vRes.ulVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Number, in kilobytes, of physical memory currently unused and available.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSFreePhysicalMemory () AS ULONGLONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT FreePhysicalMemory FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("FreePhysicalMemory", @vRes) = S_OK THEN RETURN vRes.ullVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Number, in kilobytes, that can be mapped into the operating system paging files without
' causing any other pages to be swapped out.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSFreeSpaceInPagingFiles () AS ULONGLONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT FreeSpaceInPagingFiles FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("FreeSpaceInPagingFiles", @vRes) = S_OK THEN RETURN vRes.ullVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Number, in kilobytes, of virtual memory currently unused and available.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSFreeVirtualMemory () AS ULONGLONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT FreeVirtualMemory FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("FreeVirtualMemory", @vRes) = S_OK THEN RETURN vRes.ullVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maximum number of process contexts the operating system can support. The default value
' set by the provider is 4294967295 (0xFFFFFFFF). If there is no fixed maximum, the value
' should be 0 (zero). On systems that have a fixed maximum, this object can help diagnose
' failures that occur when the maximum is reached—if unknown, enter 4294967295 (0xFFFFFFFF).
' ========================================================================================
PRIVATE FUNCTION AfxGetOSMaxNumberOfProcesses () AS ULONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT MaxNumberOfProcesses FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("MaxNumberOfProcesses", @vRes) = S_OK THEN RETURN vRes.ulVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Maximum number, in kilobytes, of memory that can be allocated to a process.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSMaxProcessMemorySize () AS ULONGLONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT MaxProcessMemorySize FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("MaxProcessMemorySize", @vRes) = S_OK THEN RETURN vRes.ullVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Number of process contexts currently loaded or running on the operating system.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSNumberOfProcesses () AS ULONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT NumberOfProcesses FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("NumberOfProcesses", @vRes) = S_OK THEN RETURN vRes.ulVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Number of user sessions for which the operating system is storing state information currently.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSNumberOfUsers () AS ULONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT NumberOfUsers FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("NumberOfUsers", @vRes) = S_OK THEN RETURN vRes.ulVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Language version of the operating system installed. The following list lists the possible
' values. Example: 2055 (&h0807) (German - Switzerland).
' ========================================================================================
PRIVATE FUNCTION AfxGetOSLanguage () AS ULONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT OSLanguage FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("OSLanguage", @vRes) = S_OK THEN RETURN vRes.ulVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Number, in kilobytes, of virtual memory. For example, this may be calculated by adding
' the amount of total RAM to the amount of paging space, that is, adding the amount of
' memory in or aggregated by the computer system to the property, SizeStoredInPagingFiles.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSTotalVirtualMemorySize () AS ULONGLONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT TotalVirtualMemorySize FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("TotalVirtualMemorySize", @vRes) = S_OK THEN RETURN vRes.ullVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Total amount, in kilobytes, of physical memory available to the operating system. This
' value does not necessarily indicate the true amount of physical memory, but what is
' reported to the operating system as available to it.
' ========================================================================================
PRIVATE FUNCTION AfxGetOSTotalVisibleMemorySize () AS ULONGLONG
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT TotalVisibleMemorySize FROM Win32_OperatingSystem")
   DIM vRes AS VARIANT
   IF pWmi.MoveNext THEN
      IF pWmi.GetVar("TotalVisibleMemorySize", @vRes) = S_OK THEN RETURN vRes.ullVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Media Access Control (MAC) address of the network adapter. A MAC address is assigned by
' the manufacturer to uniquely identify the network adapter. Example: "00:80:C7:8F:6C:96".
' ========================================================================================
PRIVATE FUNCTION AfxGetMACAddress () AS CWSTR
   DIM pWmi AS CWmiCli = "root\cimv2"
   pWmi.ExecQuery("SELECT Description FROM Win32_NetworkAdapterConfiguration")
   IF pWmi.MoveNext THEN RETURN pWmi.GetStr("Description")
END FUNCTION
' ========================================================================================
