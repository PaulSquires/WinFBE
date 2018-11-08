# CDispInvoke Class

**CDispInvoke** allows to call COM methods and properties with Free Basic.

**Include file**: CDispInvoke.inc

# Constructors

| Name       | Description |
| ---------- | ----------- |
| [Constructor(PROGID)](#Constructor1) | Creates a single uninitialized object of the class associated with a specified ProgID or CLSID. |
| [Constructor(CLSID)](#Constructor2) | Creates a single uninitialized object of the class associated with a specified CLSID. |
| [Constructor(IDispatch)](#Constructor3) | Creates a single uninitialized object of the class associated with a pointer to a Dispatch interface. |
| [Constructor(VARIANT)](#Constructor4) | Creates a single uninitialized object of the class associated with a variant of the type VT_DISPATCH. |
| [Constructor(LibName)](#Constructor5) | Loads the specified library from file and creates an instance of an object. |

# Operators

| Name       | Description |
| ---------- | ----------- |
| [Operator \*](#Operator1) | Returns the underlying IDispatch pointer. |
| [Operator LET](#Operator2) | Assigns an IDispatch pointer and increases the reference count. |
| [vptr](#vptr) | Clears the contents of the class and returns the address of the underlying IDispatch pointer. |

# Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Attach](#Attach) | Attaches a Dispatch pointer. |
| [Detach](#Detach) | Detaches a Dispatch pointer. |
| [DispInvoke](#DispInvoke) | Calls a method or a get property. |
| [DispObj](#DispObj) | Returns a counted reference of the underlying dispatch pointer. |
| [DispPtr](#DispPtr) | Returns a pointer to the dispatch interface. |
| [Get](#Get) | Calls the specified property of an interface and gets the value returned. |
| [GetArgErr](#GetArgErr) | Returns the index within rgvarg of the first argument that has an error. |
| [GetDescription](#GetDescription) | Gets the exception description. |
| [GetErrorCode](#GetErrorCode) | Returns the error code. |
| [GetHelpFile](#GetHelpFile) | Gets the fully qualified help file path. |
| [GetLastResult](#GetLastResult) | Returns the last result code returned by the last executed method. of the class. |
| [GetSource](#GetSource) | Gets the name of the exception source. |
| [GetVarResult](#GetVarResult) | Returns the last result code returne by a call to the Invoke method. |
| [GetLcid](#GetLcid) | Retrieves de locale identifier used by the class. |
| [Invoke](#Invoke) | Calls a method or a get property. |
| [SetLcid](#SetLcid) | Sets de locale identifier used by the class. |
| [Put](#Put) | Calls the specified property of an interface and sets the passed value. |
| [PutRef](#PutRef) | Assigns a value to an interface member property that contains a reference to an object. |
| [Set](#Set) | Calls the specified property of an interface and sets the passed value. |
| [SetRef](#SetRef) | Assigns a value to an interface member property that contains a reference to an object. |

# <a name="Constructor1"></a>Constructor(ProgID)

Creates a single uninitialized object of the class associated with a specified ProgID or CLSID.

```
CONSTRUCTOR CDispInvoke (BYREF wszProgID AS CONST WSTRING, BYREF wszLicKey AS WSTRING = "")
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszProgID* | The ProgID or the CLSID of the object to create.<br>A ProgID such as "MSCAL.Calendar.7"<br>A CLSID such as "{8E27C92B-1264-101C-8A2F-040224009C02}" |
| *wszLicKey* | Optional. The license key as a unicode string. |

#### Usage example

```
DIM pDispInvoke AS CDispInvoke = "Scripting.Dictionary"
' -or-
pDispInvoke = CDispInvoke(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
'    CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
```

#### Example

The following example demonstrates how to use it to create an instance of the VBScript RegExp object and execute a search.

```
#include once "Afx/CDispInvoke.inc"
USING Afx

' // Create an instance of the RegExp object
DIM pDisp AS CDispInvoke = "VBScript.RegExp"
' // To check for success, see if the value returned by the DispPtr method is not null
IF pDisp.DispPtr = NULL THEN END

' // Set some properties
' // Use VARIANT_TRUE or CTRUE, not TRUE, because Free Basic TRUE is a BOOLEAN data type, not a LONG
pDisp.Put("Pattern") = ".is"
pDisp.Put("IgnoreCase") = VARIANT_TRUE
pDisp.Put("Global") = VARIANT_TRUE

' // Execute a search
DIM pMatches AS CDispInvoke = pDisp.Invoke("Execute", "IS1 is2 IS3 is4")
' // Parse the collection of matches
IF pMatches.DispPtr THEN
   ' // Get the number of matches
   DIM nCount AS LONG = VAL(pMatches.Get("Count"))
   ' // This is equivalent to:
   ' DIM cvRes AS CVAR = pMatches.Get("Count")
   ' DIM nCount AS LONG = cvRes.ValInt
   FOR i AS LONG = 0 TO nCount -1
      ' // Get a pointer to the Match object
      ' // When using COM Automation, it's not always necessary to make sure that the
      ' // passed variant with a numeric value is of the exact type, since the standard
      ' // implementation of DispInvoke tries to coerce parameters. However, it is always
      ' // safer to use a syntax like CVAR(i, "LONG")) than CVAR(i)
'      DIM pMatch AS CDIspInvoke = pMatches.Get("Item", CVAR(i, "LONG"))   ' // or CVAR(i, "LONG"))
      DIM pMatch AS CDIspInvoke = pMatches.Get("Item", i)
      IF pMatch.DispPtr THEN
         ' // Get the value of the match and convert it to a string
         print pMatch.Get("Value")
      END IF
   NEXT
END IF

PRINT
PRINT "Press any key..."
SLEEP
```

# <a name="Constructor2"></a>Constructor(CLSID)

Creates a single uninitialized object of the class associated with a specified CLSID.

```
CONSTRUCTOR CDispInvoke (BYREF classID AS CONST CLSID)
CONSTRUCTOR CDispInvoke (BYREF wszClsID AS CONST WSTRING, BYREF wszIID AS CONST WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *classID* | The CLSID (class identifier) associated with the data and code that will be used to create the object. |
| *wszClsID* | A CLSID in string format. |
| *wszIID* | The identifier of the interface to be used to communicate with the object. |

#### Usage examples

```
DIM pDispInvoke AS CDispInvoke = CDispInvoke(CLSID_Dictionary)
' where CLSID_Dictionary has been declared as
'   DIM CLSID_Dictionary AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
```

```
DIM pDispInvoke AS CDispInvoke = (CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'    CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' and IID_IDictionary as
'    CONST IID_IDictionary = "{42C642C1-97E1-11CF-978F-00A02463E06F}"
```

# <a name="Constructor3"></a>Constructor(IDispatch)

Creates a single uninitialized object of the class associated with a pointer to a Dispatch interface.

```
CONSTRUCTOR CDispInvoke (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pdisp* | Pointer to a Dispatch interface. |
| *fAddRef* | If it is false, the object takes ownership of the interface pointer without calling AddRef. This is the usual case when we assign directly an already AddRefed pointer returned by a COM method. If it is true, then **AddRef is called**. This is needed when we pass a raw pointer. |

#### Exaample

The following example combines CDispInvoke and CWmiDisp to set the specified printer as the default one.

```
#include once "Afx/CDispInvoke.inc"
#include once "Afx/CWmiDisp.inc"
USING Afx

' // Connect with WMI in the local computer and get the properties of the specified printer
DIM pDisp AS CDispInvoke = CWmiServices( _
   $"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2:" & _
   "Win32_Printer.DeviceID='OKI B410'").ServicesObj

' // Set the printer as the default printer
pDisp.Invoke("SetDefaultPrinter")

PRINT
PRINT "Press any key..."
SLEEP
```

# <a name="Constructor4"></a>Constructor(VARIANT)

Creates a single uninitialized object of the class associated with a variant of the type VT_DISPATCH.

```
CONSTRUCTOR CDispInvoke (BYVAL vDisp AS VARIANT PTR, BYVAL fAddRef AS BOOLEAN = TRUE)
CONSTRUCTOR CDispInvoke (BYVAL vDisp AS VARIANT, BYVAL fAddRef AS BOOLEAN = TRUE)
CONSTRUCTOR CDispInvoke (BYREF cvDisp AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *vDisp* | A VT_DISPATCH variant. |
| *fAddRef* | If it is false, the object takes ownership of the interface pointer without calling AddRef. This is the usual case when we assign directly an already AddRefed pointer returned by a COM method. If it is true, then **AddRef** is called. This is needed when we pass a raw pointer. |
| *cvDisp* | A VT_DISPATCH CVAR. |

# <a name="Constructor5"></a>Constructor(LibName)

Loads the specified library from file and creates an instance of an object.

```
CONSTRUCTOR CDispInvoke (BYREF wszLibName AS CONST WSTRING, BYREF rclsid AS CONST CLSID, _
   BYREF riid AS CONST IID, BYREF wszLicKey AS WSTRING = "")
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszLibName* | Full path where the library is located. |
| *rclsid* | The CLSID (class identifier) associated with the data and code that will be used to create the object. |
| *riid* | The identifier of the interface to be used to communicate with the object. |
| *wszLicKey* | Optional. The license key. |

#### Remarks

Not every component is a suitable candidate for use under this overloaded method.
* Only in-process servers (DLLs) are supported.
* Components that are system components or part of the operating system, such as XML, Data Access, Internet Explorer, or DirectX, aren't supported
* Components that are part of an application, such Microsoft Office, aren't supported.
* Components intended for use as an add-in or a snap-in, such as an Office add-in or a control in a Web browser, aren't supported.
* Components that manage a shared physical or virtual system resource aren't supported.
* Visual ActiveX controls aren't supported because they need to be initilized and activated by the OLE container.

#### Usage example

```
DIM pDispInvoke AS CDispInvoke = (CLSID_Dictionary, IID_IDictionary)
' where CLSID_Dictionary has been declared as
'    CONST CLSID_Dictionary = "{EE09B103-97E0-11CF-978F-00A02463E06F}"
' and IID_IDictionary as
'    CONST IID_IDictionary = "{42C642C1-97E1-11CF-978F-00A02463E06F}"
```

# <a name="Operator1"></a>Operator *

Returns the underlying IDispatch pointer.

```
OPERATOR * (BYREF cdi AS CDispInvoke) AS IDispatch PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *cdi* | An instance of the CDispInvoke class. |

# <a name="Operator2"></a>Operator LET (=)

Assigns an IDispatch pointer and increases the reference count.

```
OPERATOR Let (BYVAL pdisp AS IDispatch PTR)
OPERATOR Let (BYVAL vDisp AS VARIANT PTR)
OPERATOR Let (BYVAL vDisp AS VARIANT)
OPERATOR Let (BYREF cvDisp AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pdisp* | An IDispatch pointer. |
| *vDisp* | A VT_DISPATCH variant. |
| *cvDisp* | A VT_DISPATCH CVAR. |

# <a name="vptr"></a>vptr

Clears the contents of the class and returns the address of the underlying IDispatch pointer.

```
FUNCTION vptr () AS IDispatch PTR PTR
```

#### Remarks

To pass the class to an OUT BYVAL IDispatch PTR parameter.

If we pass a **CDispInvoke** variable to a function with an OUT IDIspatch parameter without first clearing the contents of the class, we may have a memory leak.

# <a name="Attach"></a>Attach

Attaches a Dispatch pointer to the class.

```
FUNCTION Attach (BYVAL pdisp AS IDispatch PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pdisp* | Pointer to a Dispatch interface, |
| *fAddRef* | If it is false, the object takes ownership of the interface pointer without calling **AddRef**.<br>If it is true, then **AddRef** is called. This is needed when we pass a raw pointer. |

#### Return value

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |
| E_INVALIDARG | The passed pointer is null. |

# <a name="Detach"></a>Detach

Detaches the Dispatch pointer from the class.

```
FUNCTION Detach () AS IDispatch PTR
```

#### Return value

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |

#### Remarks

Extracts and returns the encapsulated interface pointer, and then clears the encapsulated pointer storage to NULL. This removes the interface pointer from encapsulation. It is up  to the caller to call **Release** on the returned interface pointer.

# <a name="DispInvoke"></a>DispInvoke

Calls a method or a get property.

```
FUNCTION DispInvoke (BYVAL wFlags AS WORD, BYVAL dispIdMember AS DISPID, BYVAL prgArgs AS VARIANT PTR = NULL, _
   BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
FUNCTION DispInvoke (BYVAL wFlags AS WORD, BYVAL pwszName AS WSTRING PTR, BYVAL prgArgs AS VARIANT PTR = NULL, _
   BYVAL cArgs AS UINT = 0, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wFlags* | Flags describing the context of the Invoke call. |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the method or property to call. |
| *prgArgs* | Array of variant parameters in reversed order. |
| *cArgs* | The number of arguments in the *prgArgs* array. |
| *lcid* | The locale context in which to interpret arguments. The lcid is used by the **GetIDsOfNames** function, and is also passed to Invoke to allow the object to interpret its arguments specific to a locale. Applications that do not support multiple national languages can ignore this parameter. |

#### Return value

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |

#### Remarks

Extracts and returns the encapsulated interface pointer, and then clears the encapsulated pointer storage to NULL. This removes the interface pointer from encapsulation. It is up  to the caller to call **Release** on the returned interface pointer.

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |
| DISP_E_BADPARAMCOUNT | The number of elements provided to **DISPPARAMS** is different from the number of arguments accepted by the method or property. |
| DISP_E_BADVARTYPE | One of the arguments in **DISPPARAMS** is not a valid variant type. |
| DISP_E_EXCEPTION | The application needs to raise an exception. In this case, the structure passed in *pexcepinfo* should be filled in. Note: Because the information that can be returned by the **EXCEPINFO** structure is very limited, some COM servers like ADO or WMI use its own system to return errors. |
| DISP_E_MEMBERNOTFOUND | The requested member does not exist. |
| DISP_E_NONAMEDARGS | This implementation of IDispatch does not support named arguments. |
| DISP_E_OVERFLOW | One of the arguments in **DISPPARAMS** could not be coerced to the specified type. |
| DISP_E_PARAMNOTFOUND | One of the parameter IDs does not correspond to a parameter on the method. In this case, puArgErr is set to the first argument that contains the error. |
| DISP_E_TYPEMISMATCH | One or more of the arguments could not be coerced. The index of the first parameter with the incorrect type within rgvarg is returned in *puArgErr*. |
| DISP_E_UNKNOWNINTERFACE | The interface identifier passed in riid is not IID_NULL. |
| DISP_E_UNKNOWNLCID | The member being invoked interprets string arguments according to the LCID, and the LCID is not recognized. If the LCID is not needed to interpret arguments, this error should not be returned. |
| DISP_E_PARAMNOTOPTIONAL | A required parameter was omitted. |

#### Remarks

This method is called internally by the **Get**, **Put**, **PutRef** and **Invoke** methods of the **CDispInvoke class**. You don't need to call it directly.

# <a name="DispObj"></a>DispObj

Returns a counted reference of the underlying dispatch pointer. You must release it, e.g. calling call **IUnknown_Release** or the function **AfxSafeRelease** when no longer need it.

```
FUNCTION DispObj () AS IDispatch PTR
```

# <a name="DispPtr"></a>DispPtr

Returns a pointer to the dispatch interface. Don't call **IUnknown_Release** on it.

```
FUNCTION DispPtr () AS IDispatch PTR
```

# <a name="Get"></a>Get

Calls the specified property of an interface and gets the value returned. The overloaded methods allow to use the **Get** method passing one argument, two arguments or none.

```
FUNCTION Get (BYVAL dispID AS DISPID) AS CVAR
FUNCTION Get (BYVAL dispID AS DISPID, BYREF cvArg AS CVAR) AS CVAR
FUNCTION Get (BYVAL dispID AS DISPID, BYREF cvArg1 AS CVAR, BYREF cvArg2 AS CVAR) AS CVAR
FUNCTION Get (BYVAL pwszName AS WSTRING PTR) AS CVAR
FUNCTION Get (BYVAL pwszName AS WSTRING PTR, BYREF cvArg AS CVAR) AS CVAR
FUNCTION Get (BYVAL pwszName AS WSTRING PTR, BYREF cvArg1 AS CVAR, BYREF cvArg2 AS CVAR) AS CVAR
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *cvArg1* | CVAR. First argument. |
| *cvArg2* | CVAR. Second argument. |

#### Return value

CVAR. The property value.

#### Example

```
#include "windows.bi"
#include "Afx/CWmiDisp.inc"
using Afx
' // Connect to WMI using a moniker
' // Note: $ is used to avoid the pedantic warning of the compiler about escape characters
DIM pServices AS CWmiServices = $"winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"
IF pServices.ServicesPtr = NULL THEN END
' // Execute a query
DIM hr AS HRESULT = pServices.ExecQuery("SELECT Caption, SerialNumber FROM Win32_BIOS")
IF hr <> S_OK THEN PRINT AfxWmiGetErrorCodeText(hr) : SLEEP : END
' // Get the number of objects retrieved
DIM nCount AS LONG = pServices.ObjectsCount
print "Count: ", nCount
IF nCount = 0 THEN PRINT "No objects found" : SLEEP : END
' // Enumerate the objects using the standard IEnumVARIANT enumerator (NextObject method)
' // and retrieve the properties using the CDispInvoke class.
DIM pDispServ AS CDispInvoke = pServices.NextObject
IF pDispServ.DispPtr THEN
   PRINT "Caption: "; pDispServ.Get("Caption")
   PRINT "Serial number: "; pDispServ.Get("SerialNumber")
END IF
PRINT
PRINT "Press any key..."
SLEEP
```

# <a name="GetArgErr"></a>GetArgErr

Returns the index within rgvarg of the first argument that has an error. Arguments are stored in pDispParams->rgvarg in reverse order, so the first argument is the one with the highest index in the array. This parameter is returned only when the resulting return value is DISP_E_TYPEMISMATCH or DISP_E_PARAMNOTFOUND.

```
FUNCTION GetArgErr () AS UINT
```

# <a name="GetDescription"></a>GetDescription

Gets the exception description.

```
FUNCTION GetDescription () AS CWSTR
```

# <a name="GetErrorCode"></a>GetErrorCode

Returns the error code. When the call to Invoke returns DISP_E_EXCEPTION, this function returns a long integer value with a more specific error code. If the value is less than 65536 it is usually an application defined code, stored in the *wCode* member of the EXCEPINFO structure. More common are the longer values, usually defined by Windows, stored in the *sCode* member, such E_INVALIDARG (&h80070057), E_FAIL (&h80004005), etc.

```
FUNCTION GetErrorCode () AS SCODE
```

# <a name="GetHelpFile"></a>GetHelpFile

Gets the fully qualified help file path. In many cases it is empty or outdated.

```
FUNCTION GetHelpFile () AS CWSTR
```

# <a name="GetLastResult"></a>GetLastResult

Returns the result code returned by the last executed method..

```
FUNCTION GetLastResult () AS HRESULT
```

# <a name="GetSource"></a>GetSource

Gets the name of the exception source. Typically, this is an application name.

```
FUNCTION GetSource () AS CWSTR
```

# <a name="GetVarResult"></a>GetVarResult

The result of a call to the Invoke method. Not usually needed because Invoke already returns it as the result of the function.

```
FUNCTION GetVarResult () AS CVAR
```

# <a name="GetLcid"></a>GetLcid

Returns de locale identifier used by the class.

```
FUNCTION GetLcid () AS LCID
```

# <a name="Invoke"></a>Invoke

Calls a method or a get property.

```
FUNCTION Invoke (BYVAL dispID AS DISPID) AS CVAR
FUNCTION Invoke (BYVAL dispID AS DISPID, cvArg1...cvArg16) AS CVAR
FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR) AS CVAR
FUNCTION Invoke (BYVAL pwszName AS WSTRING PTR, cvArg1...cvArg16) AS CVAR
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *cvArg1...cvArg16* | CVAR. Parameters that will be passed to **IDIspatch.Invoke** as an array of variants in reversed order. |

Remarks

To check for succes or failure, call the **GetLastResult** method. It will return S_OK (0) on succes or an HRESULT code on failure.

| HRESULT    | Description |
| ---------- | ----------- |
| S_OK | Success. |
| DISP_E_BADPARAMCOUNT | The number of elements provided to **DISPPARAMS** is different from the number of arguments accepted by the method or property. |
| DISP_E_BADVARTYPE | One of the arguments in **DISPPARAMS** is not a valid variant type. |
| DISP_E_EXCEPTION | The application needs to raise an exception. In this case, the structure passed in *pexcepinfo* should be filled in. Note: Because the information that can be returned by the **EXCEPINFO** structure is very limited, some COM servers like ADO or WMI use its own system to return errors. |
| DISP_E_MEMBERNOTFOUND | The requested member does not exist. |
| DISP_E_NONAMEDARGS | This implementation of IDispatch does not support named arguments. |
| DISP_E_OVERFLOW | One of the arguments in **DISPPARAMS*** could not be coerced to the specified type. |
| DISP_E_PARAMNOTFOUND | One of the parameter IDs does not correspond to a parameter on the method. In this case, *puArgErr* is set to the first argument that contains the error. |
| DISP_E_TYPEMISMATCH | One or more of the arguments could not be coerced. The index of the first parameter with the incorrect type within *rgvarg* is returned in *puArgErr*. |
| DISP_E_UNKNOWNINTERFACE | The interface identifier passed in *riid* is not IID_NULL. |
| DISP_E_UNKNOWNLCID | The member being invoked interprets string arguments according to the LCID, and the LCID is not recognized. If the LCID is not needed to interpret arguments, this error should not be returned. |
| DISP_E_PARAMNOTOPTIONAL | A required parameter was omitted. |

#### Remarks

For optional parameters, we must use a VT_ERROR VARIANT with a value of DISP_E_PARAMNOTFOUND. You can use the function **AfxCVarOptPrm** or the macro **CVAR_OPTPRM** for this purpose.

# <a name="SetLcid"></a>SetLcid

Sets de locale identifier used by the class.

```
FUNCTION SetLcid (BYVAL lcid AS LCID)
```

| Parameter  | Description |
| ---------- | ----------- |
| *lcid* | The locale context in which to interpret arguments. The *lcid* is used by the **GetIDsOfNames** function, and is also passed to **Invoke** to allow the object to interpret its arguments specific to a locale. Applications that do not support multiple national languages can ignore this parameter. |

# <a name="Put"></a>Put

Calls the specified property of an interface and sets the passed value.

```
PROPERTY Put (BYVAL dispID AS DISPID, BYREF cvArg AS CVAR)
PROPERTY Put (BYVAL pwszName AS WSTRING PTR, BYREF cvArg AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *cvArg* | CVAR. The value to set. |

# <a name="PutRef"></a>PutRef

Assigns a value to an interface member property that contains a reference to an object. For example, a reference to another interface.

```
PROPERTY PutRef (BYVAL dispID AS DISPID, BYREF cvArg AS CVAR)
PROPERTY PutRef (BYVAL pwszName AS WSTRING PTR, BYREF cvArg AS CVAR)
PROPERTY PutRef (BYVAL dispID AS DISPID, BYVAL pv AS ANY PTR)
PROPERTY PutRef (BYVAL pwszName AS WSTRING PTR, BYVAL pv AS ANY PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *cvArg* | CVAR. The value to set. |
| *pv* | Pointer to an interface. |

# <a name="Set"></a>Set

Calls the specified property of an interface and sets the passed value.

```
FUNCTION Set (BYVAL dispID AS DISPID, BYREF cvArg AS CVAR) AS HRESULT
FUNCTION Set (BYVAL dispID AS DISPID, BYREF cvArg1 ... cvArg2 AS CVAR) AS HRESULT
FUNCTION Set (BYVAL dispID AS DISPID, BYREF cvArg1 ... cvArg3 AS CVAR) AS HRESULT
FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF cvArg AS CVAR) AS HRESULT
FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF cvArg1 ... cvArg2 AS CVAR) AS HRESULT
FUNCTION Set (BYVAL pwszName AS WSTRING PTR, BYREF cvArg1 ... cvArg3 AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *cvArg* | CVAR. The value to set. |
| *cvArg1...cvArg3* | CVAR. The last parameter is the value to set. cvArg1 and cvArg2 are the index values. |

#### Returm value

S_OK (0) on success or an HRESULT code.

# <a name="SetRef"></a>SetRef

Assigns a value to an interface member property that contains a reference to an object. For example, a reference to another interface.

```
FUNCTION SetRef (BYVAL dispID AS DISPID, BYREF cvArg AS CVAR) AS HRESULT
FUNCTION SetRef (BYVAL dispID AS DISPID, BYREF cvArg1 ... cvArg2 AS CVAR) AS HRESULT
FUNCTION SetRef (BYVAL dispID AS DISPID, BYREF cvArg1 ... cvArg3 AS CVAR) AS HRESULT
FUNCTION SetRef (BYVAL pwszName AS WSTRING PTR, BYREF cvArg AS CVAR) AS HRESULT
FUNCTION SetRef (BYVAL pwszName AS WSTRING PTR, BYREF cvArg1 ... cvArg2 AS CVAR) AS HRESULT
FUNCTION SetRef (BYVAL pwszName AS WSTRING PTR, BYREF cvArg1 ... cvArg3 AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dispID* | Identifies the member. Use **GetIDsOfNames** or the object's documentation to obtain the dispatch identifier. |
| *pwszName* | The name of the property to call. |
| *cvArg1...cvArg3* | CVAR. The last parameter is the value to set. cvArg1 and cvArg2 are the index values. |

#### Returm value

S_OK (0) on success or an HRESULT code.
