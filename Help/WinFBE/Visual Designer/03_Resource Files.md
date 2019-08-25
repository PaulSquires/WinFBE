# Resource file

All WinFBE visual designer created applications need to have a resource file and a manifest file. 

If you are using WinFBE’s project management to create your program then the resource file and manifest file can be created automatically. If you are creating your application outside of a project then you need to manually provide a resource file that includes a reference to a manifest file. Failing to do this will result in your application not being theme aware. It will look ugly. 

If manually specifying a resource file then you can use WinFBE’s built-in statement to identify the resource filename. 

Don't forget the ' comment marker in front of the #RESOURCE statement

    '#RESOURCE "resource.rc"

The contents of the resource file can be as simple as the following:

```
//=====================================================================
// Generic project resource file
//=====================================================================

1 24 ".\manifest.xml"
```

The link to the manifest file is important because Windows uses it to describe the application's EXE and ensure that it activates compatibility features for the version of Windows that the EXE is being run on. 

The following is a generic manifest file that you can use:

```
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0" xmlns:asmv3="urn:schemas-microsoft-com:asm.v3">
  <assemblyIdentity version="1.0.0.0"
    processorArchitecture="*"
    name="ApplicationName"
    type="win32"/>
  <description>Optional description of your application</description>
  
  <asmv3:application>
    <asmv3:windowsSettings xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">
    <dpiAware>true</dpiAware>
    </asmv3:windowsSettings>
  </asmv3:application>
    
  <!-- Compatibility section -->
  <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
    <application>
      <!--The ID below indicates application support for Windows Vista -->
      <supportedOS Id="{e2011457-1546-43c5-a5fe-008deee3d3f0}"/>
      <!--The ID below indicates application support for Windows 7 -->
      <supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"/>
      <!--This Id value indicates the application supports Windows 8 functionality-->
      <supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"/>
      <!--This Id value indicates the application supports Windows 8.1 functionality-->
      <supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/>
      <!--This Id value indicates the application supports Windows 10 functionality-->
      <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
    </application>
  </compatibility>
  
  <!-- Trustinfo section -->
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v2">
    <security>
      <requestedPrivileges>
        <requestedExecutionLevel
          level="asInvoker"
          uiAccess="false"/>
      </requestedPrivileges>
    </security>
  </trustInfo>
    
  <dependency>
    <dependentAssembly>
      <assemblyIdentity
        type="win32"
        name="Microsoft.Windows.Common-Controls"
        version="6.0.0.0"
        processorArchitecture="*"
        publicKeyToken="6595b64144ccf1df"
        language="*" />
    </dependentAssembly>
  </dependency>
    
</assembly>
```
