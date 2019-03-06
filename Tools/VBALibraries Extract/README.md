# VBALibraries.json

This JSON file contains an exported list of Classes, Types and Modules, 
with each of their exported members for all standard Excel Libraries.

This list was generated using `VBALibraries.ahk` which parses the
object browser in the VBE.

> Note:
> The JSON files created by the tool may not be fully correct.
> We are correcting the JSONs where conflicts are found.

## JSON Structure:

```js
{
  "type": "Libraries",
  
  //All libraries
  "zChildren": [
    {
      //Description of library as given by VBA
      "LibDescription": "Microsoft Excel 16.0 Object Library",
      
      //Name of library. Note: \r\n may be virtual here
      "LibName": "Excel\r\n",
      
      //Path of library
      "LibPath": "C:\\Program Files (x86)\\Microsoft Office\\Root\\Office16\\EXCEL.EXE",
      
      //Library Type (Library or Project)
      "LibType": "Library",
      
      //Children of this library (Mostly classes and modules)
      "zChildren": [
        {
          //Name of class/module/namespace. Note: Global is a special module name denoting the global namespace
          "name": "Global",
          
          //Parent of object. Note: Global methods are always declared somewhere, and then elevated to global status.
          "parent": "Excel",
          
          //module/class - note for global this should likely be "Namespace"
          "type": "Class",
          
          //Description as provided by VBA
          "typeDescription": "<globals>",
          
          //All children of this member
          "members": [
            {
              //Is the member the default member of the object (i.e. DISPID=0 / UserMemID=0
              "isDefault": false,
              
              //Is the property or value read only?
              "isReadOnly": true,
              
              //The member name
              "MemberName": "ActiveCell",
              
              //E.G. Property, Sub, Function, ...
              "MemberType": "Property",
              
              //ParamString if given (note this might also be member default value)
              "ParamString": "",
              
              //The parent of this member
              "Parent": "Excel.Global",
              
              //Return type if given (sometimes no return value is given)
              "RetType": "Range\r\n",
              
              //Raw text exported from Excel, used to infer the above data.
              "typeDescription": "Property ActiveCell As Range\r\n    read-only\r\n    Member of Excel.Global\r\n"
            },...
          ]
        }, ...
      ]
    },...
  ]
}
```
