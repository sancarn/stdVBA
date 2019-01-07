# VBA-STD-Library

A Collection of libraries to form a common standard layer for modern VBA applications.

## Motivation

VBA first appeared in 1993 (over 25 years ago) and the language's age is definitely noticable. VBA has a lot of specific libraries for controlling Word, Excel, Powerpoint etc. However the language massively lacks in modern libraries. VBA projects ultimately become a mish mash of many different technologies. Commonly for me that means calls to Win32 DLLs, COM libraries via late-binding, calls to command line applications and calls to .NET assemblies.

Over time I have been building my own libraries and have gradually build my own layer above the simple VBA APIs.

The VBA Standard Library aims to give users a set of common libraries, maintained by the community, which aid in the building of VBA Applications.

## Planned Global Objects:

| VBType |Type       |Name             | Description  |
|--------|-----------|-----------------|--------------|
| Class  |EntryPoint |STD              | Entry point for all APIs. Mainly build for intellisense to the rest of the API.
| Class  |File       |stdFileSystem    | A wrapper around Shell's and FSO's file system APIs.
| Class  |Data       |stdDate          | A standard date parsing library. No more will you have to rely on Excel's interpreter. State the format, get the data.
| Class  |Data       |stdRegex         | A wrapper around `VBScript.RegExp` but with more features e.g. named capture groups and free-spaces.
| Class  |Data       |stdJSON          | [Tim Hall's fantastic JSON library](https://github.com/VBA-tools/VBA-JSON)
| Class  |Data       |stdHTTP          | A wrapper around HTTP COM libraries.
| Class  |Data       |stdOXML          | A library to assist in the modification of Office documents via Open XML format.
| Class  |File       |stdZip           | A wrapper around shell's Zip functionality.
| Class  |Debug      |stdDebug         | A wrapper around `Debug` while adding new options like styling messages and printing to an external html console.
| Class  |Type       |stdArray         | A library designed to re-create the Javascript dynamic array object.
| Class  |Type       |stdDictionary    | A drop in replacement for VBScript's dictionary.
| Module |Type       |stdIniVariantEnum| Initialising [IEnumVARIANT](http://www.vbforums.com/showthread.php?854963-VB6-IEnumVARIANT-For-Each-support-without-a-typelib) by recreating vtable. Used to overcome pitfalls of VB collections.
| Class  |Runtimes   |stdCLR           | Host common language runtime with VBA. This allows compilation and execution of C#.NET and VB.NET scripts in-process.
| Class  |Runtimes   |stdPowershell    | Programmatic access to Powershell from VBA.
| Class  |Runtimes   |stdJSBridge      | A VbaJsBridge module allowing applications to open and close programmatic access to VBA from OfficeJS.
| Class  |Runtimes   |stdVBR           | [Hidden functions from VB VirtualMachine library](http://www.freevbcode.com/ShowCode.asp?ID=7520)
| Class  |Runtimes   |stdExecLib       | Execute external applications in-memory. [src](https://github.com/itm4n/VBA-RunPE)
| Class  |Automation |stdAccessibility | Use Microsoft Active Accessibility framework within VBA - Very useful for automation.
| Class  |Automation |stdWindows       | Standard functions for handling Windows
| Class  |Automation |stdKernel        | Low level but useful APIs. Won't be loading Kernel32.dll entirely, but will try to expose static methods to common useful functions.
| Class  |Processing |stdThread        | Multithreading in VBA? [src](http://www.freevbcode.com/ShowCode.asp?ID=1287#A%20Quick%20Review%20Of%20Multithreading)
| Class | Application | stdEvents | More events for VBA.

## Structure

All modules or classes will be prefixed by std and will also be linked through STD class.

Commonly implementations will be of classes which are factory classes. E.G:

```vb
Class stdFile
  private initialized as boolean
  private pPath as string
  
  Public Property Let Path(sPath as value)
    if not initialized then
      pPath = sPath
    Else
      Err.Raise 0, "STD::File::Path[Setter]", "Path can only be set on an uninitialized class."
    End If
  End Property
  
  Friend Sub Initialize()
    initialized = true
  End Sub
  
  Function Open(sPath as string) as stdFile
    Set Open = new stdFile
    Open.path = sPath
    Call Open.initialize
  End Function
End Class
```

With the above example, the File class can only have it's Path property changed when the class is uninitialized. After initialized it will throw an error.
Typically the class will be initialized through `File::Open()`. We will try to keep this structure across all VBA files.
