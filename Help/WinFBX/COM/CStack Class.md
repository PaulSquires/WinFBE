# CStack Class

A Stack Collection is an ordered set of data items, which are accessed on a LIFO (Last-In / First-Out) basis. Each data item is passed and stored as a variant variable, using the Push and Pop methods.

**Include file**: CStack.inc

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Push](#Push) | Appends a variant at the end of the collection. |
| [Pop](#Pop) | Gets and removes the last element of the collection. |
| [Count](#Count) | Returns the number of elements in the collection. |
| [Clear](#Clear) | Removes all the elements in the collection. |

# <a name="Push"></a>Push

Appends a variant at the end of the collection.

```
FUNCTION Push (BYREF cv AS CVAR) AS HRESULT
```

#### Return value

Returns S_OK on success, or an error HRESULT on failure.
Error DISP_E_ARRAYISLOCKED: The array is locked.

#### Usage example:

```
#INCLUDE ONCE "Afx/CStack.inc"
using Afx

DIM pStack AS CStack
pStack.Push "String 1"
pStack.Push "String 2"
DIM cv AS CVAR
cv = pStack.Pop
PRINT cv
cv = pStack.Pop
PRINT cv

' --or--

PRINT pStack.Pop
PRINT pStack.Pop
```

# <a name="Pop"></a>Pop

Gets and removes the last element of the collection.

```
FUNCTION Pop () AS CVAR
```

#### Return value

A Variant. If there are no elements to pop, the returned variant will be empty.

# <a name="Count"></a>Count

Returns the number of elements in the collection.

```
FUNCTION Count () AS UINT
```

# <a name="Clear"></a>Clear

Returns the number of elements in the collection.

```
FUNCTION Clear () AS HRESULT
```

#### Return value

Returns S_OK on success, or an error HRESULT on failure.
