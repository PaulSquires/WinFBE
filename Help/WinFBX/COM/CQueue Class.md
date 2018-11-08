# CQueue Class

A Queue Collection is an ordered set of data items, which are accessed on a FIFO (First-In / First-Out) basis. Each data item is passed and stored as a variant variable, using the Enqueue and Dequeue methods.

**Include file**: CStack.inc

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Enqueue](#Enqueue) | Appends a variant at the end of the collection. |
| [Dequeue](#Dequeue) | Gets and removes the first element of the collection. |
| [Count](#Count) | Returns the number of elements in the collection. |
| [Clear](#Clear) | Removes all the elements in the collection. |

# <a name="Enqueue"></a>Enqueue

Appends a variant at the end of the collection.

```
FUNCTION Enqueue (BYREF cv AS CVAR) AS HRESULT
```

#### Return value

Returns S_OK on success, or an error HRESULT on failure.
Error DISP_E_ARRAYISLOCKED: The array is locked.

#### Usage example:

```
#INCLUDE ONCE "Afx/CStack.inc"
using Afx

DIM pQueue AS CQueue
pQueue.Enqueue "String 1"
pQueue.Enqueue 12345.12
print pQueue.Dequeue
print pQueue.Dequeue
```

# <a name="Dequeue"></a>Dequeue

Gets and removes the first element of the collection.

```
FUNCTION Dequeue () AS CVAR
```

#### Return value

A Variant. If there are no elements to dequeue, the returned variant will be empty.

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
