# stdArray

A dynamic list with a configurable index base. The default is 1.

## Choosing a base

Set the global default before creating arrays:

~~~vb
stdArray.DefaultBase = 0
Dim values As stdArray
Set values = stdArray.Create("a", "b")
Debug.Print values(0) 'a
Debug.Print values.Base '0
~~~

Each array snapshots the default during initialisation. Changing DefaultBase does not rebase existing arrays.

For an explicit override, use:

~~~vb
Set values = stdArray.CreateWithOptions(10, 10, 5, "a", "b")
Debug.Print values(5) 'a

'Use the current global default:
Set values = stdArray.CreateWithOptions(10, 10, Empty, "a", "b")
~~~

CreateWithOptions takes initial capacity, growth chunk size, index base, then initial items. The base argument is required because VBA cannot combine Optional parameters with ParamArray. Pass Empty for the global default. The protected protInit(capacity, chunkSize, [base]) also accepts an optional base.

## Index behavior

Item/Get/Let/Set, PutItem, Remove, Slice, Splice, search start indices and add/remove event indices use the instance base. Length is always an element count. Slice end indices are inclusive. Nonempty arr exports have the instance base; empty exports remain VBA's empty Array().

Clone, Reverse, Map, Filter, Unique, Sort, Slice, Splice's removed items, SplitToChunks and GroupBy preserve the source base. Concat and Splice accept arrays with different bases.

indexOf and lastIndexOf return -1 when absent, preserving the original contract. For a negative base, -1 may also be a valid index: use includes to check membership unambiguously.

## Migration

Ordinary Create(...) calls remain one-based until DefaultBase is changed.

CreateWithOptions has a new third argument. Replace CreateWithOptions(n, chunk, a, b) with CreateWithOptions(n, chunk, Empty, a, b), or supply the desired base explicitly. Calls with no initial items also need the third argument: CreateWithOptions(n, chunk, Empty).

