# Useful links for the COM Library

## IUnknown functions:

* [IUnknown_AtomicRelease]() - Releases a Component Object Model (COM) pointer and sets it to NULL.
* [IUnknown_GetSite]() - Calls the specified object's IObjectWithSite::GetSite method.
* [IUnknown_GetWindow]() - Attempts to retrieve a window handle from a Component Object Model (COM) object by querying for various interfaces that have a GetWindow method.
* [IUnknown_QueryService]() - Retrieves an interface for a service from a specified object.
* [IUnknown_Set]() - Changes the value of a Component Object Model (COM) interface pointer and releases the previous interface.
* [IUnknown_SetSite]() - Sets the specified object's site by calling its IObjectWithSite::SetSite method.

## `QueryInterface`

* [QISearch to find interface by IID](https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-qisearch)

...


* [SetProcessReference?](https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-setprocessreference)




## [Threading](https://docs.microsoft.com/en-us/windows/desktop/shell/managing-thread-references)

* [SHCreateThread](https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-shcreatethread)
* [SHCreateThreadRef](https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-shcreatethreadref)
* [SHCreateThreadWithHandle](https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-shcreatethreadwithhandle)

## Registry functions

...

## General

* [C String Interpolation](https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-wvnsprintfw)
* [GetProcessReference](https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-getprocessreference) - Retrieves the process-specific object supplied by SetProcessReference, incrementing the reference count to keep the process alive.


## Hashing

* [HashData](https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-hashdata) - Hashes an array of data
