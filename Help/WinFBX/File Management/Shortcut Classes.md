# Shortcut Classes

**CShortcut** allows to create a shortcut programmatically.

**CURLShortcut** allows to create a URL shortcut programmatically.

**Include file**: CShortcut.inc

### Constructors

```
CONSTRUCTOR CShortcut (BYREF cbsPathName AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsPathName* | The full path and file name of the shortcut. |

#### Example

```
' Creates a shortcut programatically (if it already exists, opens it)
DIM pShortcut AS CShortcut = ExePath & "\Test.lnk"   ' --> change me
' Sets various properties and saves them to disk
pShortcut.Description = "Hello world"   ' --> change me
pShortcut.WorkingDirectory = ExePath & "\"   ' --> change me
pShortcut.Arguments = "/c"   ' --> change me
pShortcut.HotKey = "Ctrl+Alt+e"   ' --> change me
pShortcut.IconLocation = ExePath & "\Program.ico,0"   ' --> change me
pShortcut.RelativePath = ExePath & "\"   ' --> change me
pShortcut.TargetPath = ExePath & $"\HelloWord.exe"   ' --> change me
pShortcut.WindowStyle = WshNormalFocus
pShortcut.Save
```

```
CONSTRUCTOR CURLShortcut (BYREF cbsPathName AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsPathName* | The full path and file name of the shortcut. |

```
' Creates a shortcut programatically (if it already exists, opens it)
DIM pURLShortcut AS CURLShortcut = ExePath & "\Microsoft Web Site.url"   ' --> change me
pURLShortcut.TargetPath = "http://www.microsoft.com"   ' --> change me
pURLShortcut.Save
```

# CShortcut Methods

| Name       | Description |
| ---------- | ----------- |
| [Save](#Save1) | Saves a shortcut object to disk. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |

# CShortcut Properties

| Name       | Description |
| ---------- | ----------- |
| [Arguments](#Arguments) | Gets/sets the arguments for a shortcut, or identifies a shortcut's arguments. |
| [Description](#Description) | Returns or sets a shortcut's description. |
| [FullName](#FullName1) | Returns the fully qualified path of the shortcut object's target. |
| [Hotkey](#Hotkey) | Assigns a key-combination to a shortcut, or identifies the key-combination assigned to a shortcut. |
| [IconLocation](#IconLocation) | Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut. |
| [RelativePath](#RelativePath) | Assigns a relative path to a shortcut. |
| [TargetPath](#TargetPath1) | Gets/sets the path of the shortcut's executable. |
| [WindowStyle](#WindowStyle) | Assigns a window style to a shortcut, or identifies the type of window style used by a shortcut. |
| [WorkingDirectory](#WorkingDirectory) | Assigns a working directory to a shortcut, or identifies the working directory used by a shortcut. |

# CURLShortcut Methods

| Name       | Description |
| ---------- | ----------- |
| [Save](#Save2) | Saves a URL shortcut object to disk. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |

# CShortcut Properties

| Name       | Description |
| ---------- | ----------- |
| [FullName](#FullName2) | Returns the fully qualified path of the shortcut object's target. |
| [TargetPath](#TargetPath2) | Gets/sets the path of the shortcut's executable. |

# <a name="Save1"></a>Save (CShortcut)

Saves a shortcut object to disk.

```
FUNCTION Save () AS HRESULT
```

#### Remarks

After creating an instance of the *CShortcut* class to create to create a shortcut object and set the shortcut object's properties, the Save method must be used to save the shortcut object to disk. The **Save** method uses the information in the shortcut object's **FullName** property to determine where to save the shortcut object on a disk. You can only create shortcuts to system objects. This includes files, directories, and drives (but does not include printer links or scheduled tasks).

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="Save2"></a>Save (CURLShortcut)

Saves a URL shortcut object to disk.

```
FUNCTION Save () AS HRESULT
```

#### Remarks

After creating an instance of the *CURLShortcut* class to create to create a shortcut object and set the shortcut object's properties, the Save method must be used to save the shortcut object to disk. The **Save** method uses the information in the shortcut object's **FullName** property to determine where to save the shortcut object on a disk. You can only create shortcuts to system objects. This includes files, directories, and drives (but does not include printer links or scheduled tasks).

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

# <a name="Arguments"></a>Arguments

Gets/sets the arguments for a shortcut, or identifies a shortcut's arguments.

```
PROPERTY Arguments () AS CBSTR
PROPERTY Arguments (BYREF cbsArguments AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsArguments* | The arguments for the shortcut. |

#### Return value

CBSTR. The arguments of the shortcut.

# <a name="Description"></a>Description

Returns or sets a shortcut's description.

```
PROPERTY Description () AS CBSTR
PROPERTY Description (BYREF cbsDescription AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsDescription* | A string value describing a shortcut. |

#### Return value

CBSTR. The description of the shortcut.

# <a name="FullName1"></a>FullName (CShortcut)

Returns the fully qualified path of the shortcut object's target.

```
PROPERTY FullName () AS CBSTR
```

#### Remarks

The **FullName** property contains a read-only string value indicating the fully qualified path to the shortcut's target.

# <a name="FullName2"></a>FullName (CURLShortcut)

Returns the fully qualified path of the shortcut object's target.

```
PROPERTY FullName () AS CBSTR
```

#### Remarks

The **FullName** property contains a read-only string value indicating the fully qualified path to the shortcut's target.

# <a name="Hotkey"></a>Hotkey

Assigns a key-combination to a shortcut, or identifies the key-combination assigned to a shortcut.

```
PROPERTY Hotkey () AS CBSTR
PROPERTY Hotkey (BYREF cbsHotkey AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsHotkey* | A string representing the key-combination to assign to the shortcut. |

#### Return value

CBSTR. The hotkey of the shortcut.

# <a name="IconLocation"></a>IconLocation

Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut.

```
PROPERTY IconLocation () AS CBSTR
PROPERTY IconLocation (BYREF cbsIconLocation AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsIconLocation* | A string that locates the icon. The string should contain a fully qualified path and an index associated with the icon. |

#### Return value

CBSTR. The location of the icon.

# <a name="RelativePath"></a>RelativePath

Assigns a relative path to a shortcut.

```
PROPERTY RelativePath (BYREF cbsRelativePath AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsRelativePath* | A string containing the relative path. |

# <a name="TargetPath1"></a>TargetPath (CShortcut)

Gets/sets the path of the shortcut's executable.

```
PROPERTY TargetPath () AS CBSTR
PROPERTY TargetPath (BYREF cbsTargetPath AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsTargetPath* | A string containing the path of the shortcut's executable. |

#### Return value

CBSTR. The path of the shortcut's executable.

# <a name="TargetPath2"></a>TargetPath (CURLShortcut)

Gets/sets the path of the shortcut's executable.

```
PROPERTY TargetPath () AS CBSTR
PROPERTY TargetPath (BYREF cbsTargetPath AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsTargetPath* | A string containing the path of the shortcut's executable. |

#### Return value

CBSTR. The path of the shortcut's executable.

# <a name="WindowStyle"></a>WindowStyle

Assigns a window style to a shortcut, or identifies the type of window style used by a shortcut.

```
PROPERTY WindowStyle () AS INT_
PROPERTY WindowStyle (BYVAL nWindowStyle AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWindowStyle* | The window style for the program being run. |

The following table lists the available settings for nWindowStyle.

| Style      | Description |
| ---------- | ----------- |
| 1 | Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. |
| 3 | Activates the window and displays it as a maximized window. |
| 7 | Minimizes the window and activates the next top-level window. |

#### Return value

LONG. The window style.

# <a name="WorkingDirectory"></a>WorkingDirectory

Assigns a working directory to a shortcut, or identifies the working directory used by a shortcut.

```
PROPERTY WorkingDirectory () AS CBSTR
PROPERTY WorkingDirectory (BYREF cbsWorkingDirectory AS CBSTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cbsWorkingDirectory* | Directory in which the shortcut starts. |

#### Return value

CBSTR. The working directory of the shortcut.
