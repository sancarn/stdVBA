var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));

// main.ts
var fs = __toESM(require("fs"));
function parseParameters(params, udtInfo) {
  let paramExtractor = /(?<optional>optional\s+)?(?:(?<referenceType>byval|byref)\s+)?(?:(?<paramarray>paramarray)\s+)?(?<name>\w+)(?<isArray>\(\))?(?:\s+as\s+(?<type>[^, )]+))?(?:\s*=\s*(?<defaultValue>.+))?/i;
  let aParams = params.split(",").map((param) => param.trim().match(paramExtractor)).map((match) => match?.groups);
  if (!aParams || aParams.length === 0) return [];
  return aParams.map((param) => {
    if (!param) return null;
    const isUDTParamType = udtInfo.some((udt) => udt.name.toLowerCase() === param.type?.toLowerCase());
    return {
      name: param.name,
      type: param.type || "",
      referenceType: param.referenceType || "",
      isOptional: !!param.optional,
      defaultValue: param.defaultValue ? param.defaultValue.trim() : void 0,
      isParamArray: !!param.paramarray,
      isArray: !!param.isArray,
      isUDTParamType
    };
  }).filter((param) => param !== null);
}
function main() {
  let files = fs.readdirSync(__dirname + "/../../src");
  files = files.filter(
    (f) => fs.lstatSync(__dirname + "/../../src/" + f).isFile()
  );
  for (const file of files) {
    let content = fs.readFileSync(__dirname + "/../../src/" + file, "utf8");
    const moduleNameFinder = /Attribute VB_Name = "(?<name>[^"]+)"/i;
    const moduleName = moduleNameFinder.exec(content)?.groups?.name ?? file.split(".")[0];
    content = content.replace(/Err\.Raise/g, "Err_Raise");
    content = content.replace(/On Error GoTo 0/g, "On Error GoTo stdErrorWrapper_ErrorOccurred");
    const udtFinder = /(?<!').*\bType\s+(?<name>\w+)/gim;
    const udtInfo = Array.from(content.matchAll(udtFinder)).map((match) => {
      return {
        name: match.groups?.name || ""
      };
    });
    const functionFinder = /(?<header>(?<!')(?:Public|Private|Friend) (?:(?<type>Function|Sub|Property) ?(?<access>Get|Let|Set)?) (?<name>\w+)\((?<params>(?:\(\)|[^)])*)\)(?: as (?<retType>(?:\w+\.)?\w+))?)(?<body>(?:.|\s)+?)\b(?<footer>End\s+(?:Function|Sub|Property))/gim;
    content = content.replace(functionFinder, (match, header, type, access, name, params, retType, body, footer, offset, haystack, groups) => {
      const conditionalCompilation = /(?<!')(?:Public|Private|Friend) (?:(?<type>Function|Sub|Property) ?(?<access>Get|Let|Set)?) (?<name>\w+)\((?<params>(?:\(\)|[^)])*)\)(?: as (?<retType>(?:\w+\.)?\w+))?(?:.|\s)+?#End If/gim;
      const conditionalCompilationMatch = conditionalCompilation.exec(body);
      if (!!conditionalCompilationMatch) {
        header = header + body.substring(0, conditionalCompilationMatch.index + conditionalCompilationMatch[0].length);
        body = body.substring(conditionalCompilationMatch.index + conditionalCompilationMatch[0].length);
      }
      let callstackName = moduleName + "#" + name + (!!access ? "[" + access + "]" : "");
      const paramsInfo = parseParameters(params, udtInfo);
      const finalParams = paramsInfo.filter((p) => !p.isUDTParamType && !p.isParamArray && !p.isArray).map((p) => `"${p.name}", ${p.name}`);
      const paramsString = (finalParams.length > 0 ? ", " : "") + finalParams.join(", ");
      const injectorHeader = [
        `  With stdError.getSentry("${callstackName}"${paramsString})`,
        "    On Error GoTo stdErrorWrapper_ErrorOccurred"
      ].join("\r\n");
      const injectorFooter = [
        "    Exit " + type,
        "    stdErrorWrapper_ErrorOccurred:",
        "      Call Err_Raise(Err.Number, Err.Source, Err.Description)",
        "  End With"
      ].join("\r\n");
      body = body.split("\n").map((line) => "    " + line).join("\n");
      return `${header}\r
${injectorHeader}\r
${body}\r
${injectorFooter}\r
${footer}`;
    });
    content += `


Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
`;
    fs.writeFileSync(__dirname + "/../../src/" + file, content, "utf8");
  }
}
main();
//# sourceMappingURL=main.js.map
