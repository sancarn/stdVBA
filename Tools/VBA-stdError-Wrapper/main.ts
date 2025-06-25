//Process description:

/* 
 1. Read all files in ../../src directory
 2. Append a function to the end of each files as follows:

```vb
Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
```
3. Replace all calls to `Err.Raise` to calls to `Err_Raise`
4. Replace all calls to `On Error GoTo 0` with `On Error GoTo stdErrorWrapper_ErrorOccurred`
5. Perform the following mapping:

```vb
Public Function MyMethod(ByVal param1 as type1, ByVal param2 as type2, ...) as returnType
  ...
End Function
```

to

```vb
Public Function MyMethod(ByVal param1 as type1, ByVal param2 as type2, ...) as returnType
  With stdError.getSentry("ModuleName#MMyethod (Access)", "param1", param1, "param2", param2, ...)
  On Error GoTo stdErrorWrapper_ErrorOccurred
    ...
  stdErrorWrapper_ErrorOccurred:
    Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function
```

6. Save and overwrite the files

This will be run on a every commit, and will publish to the `stdErrorWrapped` branch.

Watchouts:
---------------------------------------
1. Conditional compilation directives should be preserved

```vb
#If VBA7 Then
  Public Function MyMethod(...)
#Else
  Public Function MyMethod(...)
#End If
  ...
End Function
```
----------------------------------------

*/


type IUDTInfo = {
  name: string
}
type IParameter = {
  name: string;
  type: string;
  referenceType: string;
  isOptional: boolean;
  defaultValue?: any;
  isParamArray: boolean;
  isArray: boolean;
  isUDTParamType: boolean;
};

/**
 * obtain information about the parameters declared by the function, subroutine or property
 * @param params a parameter string e.g. ``ByVal sClassName As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, Optional ByVal sCaption As String = vbNullString, Optional ByVal dwStyle As Long = WS_POPUP`
 * @param udtInfo an array of IUDTInfo objects which contain information about user defined types defined in the class/moudule
 * @returns an array of IParameter objects which contain information about the parameters
 */
function parseParameters(params: string, udtInfo: IUDTInfo[]): IParameter[] {
  let paramExtractor = /(?<optional>optional\s+)?(?:(?<referenceType>byval|byref)\s+)?(?:(?<paramarray>paramarray)\s+)?(?<name>\w+)(?<isArray>\(\))?(?:\s+as\s+(?<type>[^, )]+))?(?:\s*=\s*(?<defaultValue>.+))?/i
  let aParams = params.split(",").map((param)=>param.trim().match(paramExtractor)).map((match)=>match?.groups)
  if (!aParams || aParams.length === 0) return [];

  return aParams.map((param) => {
    if (!param) return null;
    const isUDTParamType = udtInfo.some(udt => udt.name.toLowerCase() === param.type?.toLowerCase());
    return {
      name: param.name,
      type: param.type || "",
      referenceType: param.referenceType || "",
      isOptional: !!param.optional,
      defaultValue: param.defaultValue ? param.defaultValue.trim() : undefined,
      isParamArray: !!param.paramarray,
      isArray: !!param.isArray,
      isUDTParamType
    };
  }).filter(param => param !== null);
} 

import { throws } from "assert";
import * as fs from "fs";
function main() {
  //Find all files in ../../src directory
  let files = fs.readdirSync(__dirname + "/../../src");
  files = files.filter((f) =>
    fs.lstatSync(__dirname + "/../../src/" + f).isFile()
  );

  //Loop through each file
  for (const file of files) {
    //Read the file
    let content = fs.readFileSync(__dirname + "/../../src/" + file, "utf8");

    //Find the module name
    const moduleNameFinder = /Attribute VB_Name = "(?<name>[^"]+)"/i;
    const moduleName =
      moduleNameFinder.exec(content)?.groups?.name ?? file.split(".")[0];

    //Replace all calls to `Err.Raise` to calls to `Err_Raise`
    content = content.replace(/Err\.Raise/g, "Err_Raise");
    content = content.replace(/On Error GoTo 0/g, "On Error GoTo stdErrorWrapper_ErrorOccurred");

    //Get all UDTs defined in the file
    const udtFinder = /(?<!').*\bType\s+(?<name>\w+)/gim;
    const udtInfo: IUDTInfo[] = Array.from(content.matchAll(udtFinder)).map((match) => {
      return {
        name: match.groups?.name || ""
      };
    });

    //Loop through each public function
    const functionFinder = /(?<header>(?<!')(?:Public|Private|Friend) (?:(?<type>Function|Sub|Property) ?(?<access>Get|Let|Set)?) (?<name>\w+)\((?<params>(?:\(\)|[^)])*)\)(?: as (?<retType>(?:\w+\.)?\w+))?)(?<body>(?:.|\s)+?)\b(?<footer>End\s+(?:Function|Sub|Property))/gim
    content = content.replace(functionFinder, (match: string, header: string, type: string, access: string, name: string, params: string, retType: string, body: string, footer: string, offset: number, haystack: string, groups: any): string => {
      //Check if the body has another declare in it followed by a `#End If` declaration.
      const conditionalCompilation = /(?<!')(?:Public|Private|Friend) (?:(?<type>Function|Sub|Property) ?(?<access>Get|Let|Set)?) (?<name>\w+)\((?<params>(?:\(\)|[^)])*)\)(?: as (?<retType>(?:\w+\.)?\w+))?(?:.|\s)+?#End If/gim
      
      //Redefine body and header to include / exclude the conditional compilation directives as needed
      const conditionalCompilationMatch = conditionalCompilation.exec(body);
      if (!!conditionalCompilationMatch) {
        header = header + body.substring(0,conditionalCompilationMatch.index + conditionalCompilationMatch[0].length)
        body = body.substring(conditionalCompilationMatch.index + conditionalCompilationMatch[0].length)
      }

      //Get the callstack name
      let callstackName = moduleName + "#" + name + ((!!access) ? "[" + access + "]" : "");

      //Parse the parameters
      const paramsInfo = parseParameters(params, udtInfo);

      //TODO: Handle UDTs
      //TODO: Handle ParamArray
      //TODO: Handle Arrays
      const paramsString = paramsInfo.filter(p => !p.isUDTParamType && !p.isParamArray && !p.isArray).map(p => `"${p.name}", ${p.name}`).join(", ");

      const injectorHeader = [
        `  With stdError.getSentry("${callstackName}", ${paramsString})`,
        "    On Error GoTo stdErrorWrapper_ErrorOccurred"
      ].join("\r\n");

      const injectorFooter = [
        "  stdErrorWrapper_ErrorOccurred:",
        "    Call Err_Raise(Err.Number, Err.Source, Err.Description)",
        "  End With"
      ].join("\r\n");

      //Indent all lines of body by 4 spaces
      body = body.split("\n").map(line => "    " + line).join("\n");

      return `${header}\r\n${injectorHeader}\r\n${body}\r\n${injectorFooter}\r\n${footer}`;
    })

    //Append the function to the end of the file
    content += `\n\n
Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
`;

    //Save the file
    fs.writeFileSync(__dirname + "/../../src/" + file, content, "utf8");
  }
}

main();
