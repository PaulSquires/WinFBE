# CDicObj Class

**CDicObj** is an associative array of variants. Each item is associated with a unique key. The key is used to retrieve an individual item.

#### Example

```
#define unicode
#INCLUDE ONCE "Afx/CDicObj.inc"
USING Afx

' // Creates an instance of the CDicObj class
DIM pDic AS CDicObj

' // Adds some key, value pairs
pDic.Add "a", "Athens"
pDic.Add "b", "Belgrade"
pDic.Add "c", "Cairo"

print "Count: "; pDic.Count
print pDic.Exists("a")
print

' // Retrieve an item and display it
print pDic.Item("b")
print

' // Change key "b" to "m" and "Belgrade" to "México"
pDic.Key("b") = "m"
pDic.Item("m") = "México"
print pDic.Item("m")
print

' // Get all the items and display them
DIM cvItems AS CVAR = pDic.Items
FOR i AS LONG = cvItems.GetLBound TO cvItems.GetUBound
  print cvItems.GetVariantElem(i)
NEXT
print

' // Get all the keys and display them
DIM cvKeys AS CVAR = pDic.Keys
FOR i AS LONG = cvKeys.GetLBound TO cvKeys.GetUBound
  print cvKeys.GetVariantElem(i)
NEXT
print

' // Remove key "m"
pDic.Remove "m"
IF pDic.Exists("m") THEN PRINT "Key m exists" ELSE PRINT "Key m doesn't exists"

' // Remove all keys
pDic.RemoveAll
print "All the keys must have been deleted"
print "Count: "; pDic.Count
print

PRINT "Press any key..."
SLEEP
```

**Include file**: CDicObj.inc

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Add](#Add) | Adds a key and item pair to the associtive array. |
| [Count](#Count) | Returns the number of items in the associative array. |
| [DispPtr](#DispPtr) | Returns the underlying dispatch pointer. |
| [Exists](#Exists) | Checks if a specified key exists in the associative array. |
| [GetLastResult](#GetLastResult) | Returns the last result code. |
| [HashVal](#HashVal) | Returns the hash value for a specified key in the associative array. |
| [Item](#Item) | Sets or returns an item for a specified key in associative array. |
| [Items](#Items) | Returns a safe array containing all the items in the associative array. |
| [Key](#Key) | Sets or returns an item for a specified key in the associative array. |
| [Keys](#Keys) | Returns an array containing all the keys in the associative array. |
| [NewEnum](#NewEnum) | Returns a reference to the standard enumerator. |
| [Remove](#Remove) | FUNCTION Remove (BYREF cvKey AS CVAR) AS HRESULT |
| [RemoveAll](#RemoveAll) | Removes all key, item pairs from the associative array. |

# <a name="Add"></a>Add

Adds a key and item pair to the associtive array.

```
FUNCTION Add (BYREF cbKey AS CVAR, BYREF cvItem AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvKey* | The key associated with the item being added. |
| *cvItem* | The item associated with the key being added. |

#### Return value

An error occurs if the key already exists.

#### Remark

As it is a raw pointer, don't call IUnknown_Release on it.

# <a name="Count"></a>Count

Returns the number of items in the associative array.

```
FUNCTION Count () AS LONG
```

# <a name="DispPtr"></a>DispPtr

Returns the underlying dispatch pointer.

```
FUNCTION DispPtr () AS ANY PTR
```

# <a name="Exists"></a>Exists

Checks if a specified key exists in the associative array.

```
FUNCTION Exists (BYREF cvKey AS CVAR) AS BOOLEAN
```

#### Return value

# <a name="GetLastResult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

# <a name="HashVal"></a>HashVal

Returns the hash value for a specified key in the associative array.

```
FUNCTION HashVal (BYREF cvKey AS CVAR) AS CVAR
```

# <a name="Item"></a>Item

Sets or returns an item for a specified key in associative array.

```
PROPERTY Item (BYREF cvKey AS CVAR) AS CVAR it returns
PROPERTY Item (BYREF cvKey AS CVAR, BYREF cvNewItem AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvKey* | Key associated with the item being retrieved or added. |
| *cvNewvItem* | The new value associated with the specified key. |

#### Return value

The item value.

#### Remarks

If key is not found when changing an item, a new key is created with the specified *cvNewvItem*. Contrarily to the Dictionary object, if key is not found when attempting to return an existing item, it returns and empty variant and sets the last result code to E_INVALIDARG, instead of creating a new key with its corresponding item empty.

# <a name="Items"></a>Items

Returns a safe array containing all the items in the associative array.

```
FUNCTION Items () AS CVAR
```

#### Return value

A CVAR containing all the items in a safe array.

# <a name="Key"></a>Key

Sets or returns an item for a specified key in the associative array.

```
PROPERTY Key (BYREF cvKey AS CVAR, BYREF cvNewKey AS CVAR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvKey* | Key value being changed. |
| *cvNewKey* | New value that replaces the specified key. |

#### Remarks

Contrarily to the Dictionary object, if key is not found when changing a key, this method sets the last result code to E_INVALIDARG and exits, instead of creating a new key with its associated item empty.

#### Return value

The item value.

# <a name="Keys"></a>Keys

Returns an array containing all the keys in the associative array.

```
FUNCTION Keys () AS CVAR
```

#### Return value

A CVAR containing all the keys in a safe array.

# <a name="NewEnum"></a>NewEnum

Returns a reference to the standard enumerator.

```
FUNCTION NewEnum () AS IEnumVARIANT PTR
```

#### Return value

A pointer to the standard IEnumVARIANT interface.

#### Return value

IUnknown pointer that must be cast to an **IEnumVARIANT** interface.

# <a name="Remove"></a>Remove

Removes a key, item pair from the associative array.

```
FUNCTION Remove (BYREF cvKey AS CVAR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cvKey* | Key associated with the key, item pair you want to remove from the associative array. |

#### Return value

An error occurs if the specified key, item pair does not exist.

# <a name="RemoveAll"></a>RemoveAll

Removes all key, item pairs from the associative array.

```
FUNCTION RemoveAll() AS HRESULT
```

#### Return value

Returns S_OK (0) or an HRESULT error code.
