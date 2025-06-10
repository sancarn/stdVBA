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
function log(message, type = "info") {
  switch (type) {
    case "info":
      console.log(`\x1B[36m\u2139\uFE0F  Info: ${message}\x1B[0m`);
      break;
    case "warn":
      console.log(`\x1B[33;1m\u26A0\uFE0F  Warn: ${message}\x1B[0m`);
      break;
    case "error":
      console.log(`\x1B[31;1m\u274C  Error: ${message}\x1B[0m`);
      break;
    case "success":
      console.log(`\x1B[32;1m\u2705  Success: ${message}\x1B[0m`);
      break;
  }
}
function groupBy(list, keyGetter) {
  const map = /* @__PURE__ */ Object.create(null);
  if (!list) return map;
  list.forEach((item) => {
    const key = keyGetter(item);
    const collection = map[key];
    if (!collection) {
      map[key] = [item];
    } else {
      collection.push(item);
    }
  });
  return map;
}
function parseToTagLines(comment) {
  const groupByRx = /'@(?<type>\w+)(?<content>.*\s+(?:'[^@][^\n]*\s+)*)?/g;
  const matches = [...comment.matchAll(groupByRx)];
  return matches.map((match) => ({
    tag: match.groups?.type,
    content: match.groups?.content?.replace(/^'/gm, "") ?? ""
  }));
}
function parseComment(comment) {
  if (!comment) return [];
  if (!/@description/.test(comment))
    comment = comment.replace(/^'/g, "'@description ");
  const tagLines = parseToTagLines(comment);
  const regexTags = {
    description: /^\s*(?<description>[\s\S]+)/i,
    param: /^\s*(?<name>\w+)\s*(?:as\s+(?<type>[^-]+))?(?:\s*-\s*(?<description>[\s\S]+))?$/i,
    //regex needs work
    returns: /^\s*(?<type>[^-\r\n]+)?(?:\s*-\s*(?<description>[\s\S]+))?$/i,
    constructor: /(?:constructor)?/g,
    //overwrites native constructor
    throws: /(?<errNumber>\d+)\s*,\s*(?<errText>.+)/i
  };
  const commentStore = [];
  for (let tagLine of tagLines) {
    let tag = tagLine.tag;
    let groups;
    if (!!regexTags[tag]) {
      groups = regexTags[tag].exec(tagLine.content)?.groups;
    }
    switch (tag) {
      case "description":
        commentStore.push({ tag, data: tagLine.content.trim() });
        break;
      case "param":
        if (!!groups?.name)
          commentStore.push({
            tag,
            data: {
              name: groups.name.trim(),
              type: groups?.type?.trim(),
              description: groups?.description.trim()
            }
          });
        break;
      case "returns":
        commentStore.push({
          tag,
          data: {
            type: groups?.type?.trim(),
            description: groups?.description.trim()
          }
        });
        break;
      case "example":
        commentStore.push({ tag, data: tagLine.content.trim() });
        break;
      case "remark":
        commentStore.push({ tag, data: tagLine.content.trim() });
        break;
      case "devNote":
        commentStore.push({ tag, data: tagLine.content.trim() });
        break;
      case "constructor":
        commentStore.push({ tag });
        break;
      case "protected":
        commentStore.push({ tag });
        break;
      case "deprecated":
        commentStore.push({ tag, data: tagLine.content.trim() });
        break;
      case "TODO":
        commentStore.push({ tag, data: tagLine.content.trim() });
        break;
      case "throws":
        commentStore.push({
          tag,
          data: {
            errNumber: Number(groups?.errNumber),
            errText: groups?.errText
          }
        });
        break;
      case "requires":
        commentStore.push({ tag, data: tagLine.content.trim() });
        break;
      case "static":
        commentStore.push({ tag });
        break;
      default:
        log(`Unknown tag "${tag}"`, "warn");
    }
  }
  return commentStore;
}
function parseParams(params, dataParams) {
  const paramData = groupBy(
    dataParams,
    (d) => d.data.name.toLowerCase()
  );
  const paramRegex = /(?<optional>optional\s+)?(?:(?<referenceType>byval|byref)\s+)?(?:(?<paramarray>paramarray)\s+)?(?<name>\w+)(?<isArray>\(\))?(?:\s+as\s+(?<type>[^, )]+))?(?:\s*=\s*(?<defaultValue>[^,\)]+))?/gi;
  const paramMatches = [...params.matchAll(paramRegex)];
  const paramStore = [];
  for (let paramMatch of paramMatches) {
    let name = paramMatch.groups?.name;
    if (!name) continue;
    if (!paramData[name.toLowerCase()]?.length) continue;
    const commentData = paramData[name.toLowerCase()][0].data;
    let paramType = commentData.type ?? paramMatch.groups?.type ?? "Variant";
    let paramDescription = commentData.description ?? "";
    let paramReferenceType = paramMatch.groups?.referenceType ?? "ByRef";
    let paramIsParamArray = !!paramMatch.groups?.paramArray;
    let paramIsArray = !!paramMatch.groups?.isArray;
    let paramIsOptional = !!paramMatch.groups?.optional;
    let paramDefaultValue = !!paramMatch.groups?.optional ? paramMatch.groups.defaultValue ?? null : null;
    if (paramIsArray && !!commentData.type) paramType = `Array<${paramType}>`;
    paramStore.push({
      tag: "param",
      data: {
        name: name.trim(),
        type: paramType.trim(),
        description: paramDescription.trim(),
        referenceType: paramReferenceType,
        paramArray: paramIsParamArray,
        optional: paramIsOptional,
        defaultValue: paramDefaultValue
      }
    });
  }
  return paramStore;
}
function parseModuleOrClass(content, fileName) {
  let isClass = /^VERSION 1.0 CLASS/.test(content);
  let regexConditionalCompilation = /#if.+then\s+((.|\s)+?)#end if/gi;
  while (regexConditionalCompilation.test(content)) {
    content = content.replace(regexConditionalCompilation, "$1");
  }
  const moduleNameFinder = /Attribute VB_Name = "(?<name>[^"]+)"/i;
  const moduleName = moduleNameFinder.exec(content)?.groups?.name ?? fileName.split(".")[0];
  log(`Parsing module "${moduleName}"`);
  const moduleDocsFinder = /^'@module\r?\n(?:'.*\r?\n?)*/im;
  const moduleDocsString = moduleDocsFinder.exec(content)?.[0];
  const moduleDocs = parseComment(moduleDocsString);
  const moduleDocsByTag = groupBy(moduleDocs, (c) => c.tag);
  const moduleTODOs = Array.from(content.matchAll(/'TODO: (.*)/gi)).map(
    (m) => m[1]
  );
  const defaultMemberFinder = /Attribute (\w+).VB_(Var)?UserMemId += +0/i;
  const defaultMember = defaultMemberFinder.exec(content)?.groups?.[1];
  const implementsFinder = /^implements +(?<interface>\w+)/gi;
  const _implements = Array.from(content.matchAll(implementsFinder)).map(
    (m) => m.groups?.interface
  );
  const memberDocsFinder = /(?<comments>(?:\'.*\r?\n)*)(?<!' *)Public (?:(?<type>Function|Sub|Event|Property) ?(?<access>Get|Let|Set)?) (?<name>\w+)\((?<params>[^)]*)\)(?: as (?<retType>\w+))?/gim;
  const memberMatches = Array.from(content.matchAll(memberDocsFinder));
  let constructors = [];
  let events = [];
  let properties = [];
  let methods = [];
  let membersByName = groupBy(
    memberMatches,
    (m) => m.groups?.name.toLowerCase()
  );
  let memberAlreadyProcessed = {};
  memberMatches.forEach((match) => {
    let sComment = match.groups?.comments;
    let sType = match.groups?.type.toLowerCase();
    let sName = match.groups?.name;
    let sParams = match.groups?.params ?? "";
    let sRetType = match.groups?.retType ?? "Void";
    if (memberAlreadyProcessed[sName.toLowerCase()]) return;
    memberAlreadyProcessed[sName.toLowerCase()] = true;
    let access;
    if (sType === "property") {
      let accessTypes = membersByName[sName.toLowerCase()].map(
        (match2) => match2.groups?.access.toLowerCase()
      );
      let readAccess = accessTypes.includes("get");
      let writeAccess = accessTypes.includes("let") || accessTypes.includes("set");
      if (readAccess && writeAccess) {
        access = "ReadWrite";
      } else if (readAccess) {
        access = "ReadOnly";
      } else if (writeAccess) {
        access = "WriteOnly";
      }
    }
    let commentData;
    if (sComment.length > 0) {
      commentData = parseComment(sComment);
    }
    let commentDataByTag = groupBy(commentData, (c) => c.tag);
    let params = parseParams(
      sParams,
      commentDataByTag["param"]
    ).map((param) => {
      return {
        name: param.data.name,
        type: param.data.type,
        description: param.data?.description ?? "",
        optional: param.data?.optional ?? false,
        defaultValue: param.data?.defaultValue ?? null,
        paramArray: param.data?.paramArray ?? false,
        referenceType: param.data?.referenceType ?? "ByRef"
      };
    });
    switch (sType) {
      case "sub":
      case "function":
        let arrToPushTo = !!commentDataByTag["constructor"]?.length ? constructors : methods;
        let func = {
          name: sName,
          description: commentDataByTag["description"]?.[0]?.data ?? "",
          remarks: commentDataByTag["remark"]?.map((c) => c.data) ?? [],
          examples: commentDataByTag["example"]?.map((c) => c.data) ?? [],
          params,
          returns: sType === "sub" ? null : {
            type: commentDataByTag["returns"]?.[0]?.data.type ?? sRetType,
            description: commentDataByTag["returns"]?.[0]?.data.description ?? ""
          },
          deprecation: {
            status: !!commentDataByTag["deprecated"]?.length,
            message: commentDataByTag["deprecated"]?.[0]?.data ?? ""
          },
          isDefaultMember: defaultMember === sName,
          devNotes: commentDataByTag["devNote"]?.map((c) => c.data) ?? [],
          todos: commentDataByTag["todo"]?.map((c) => c.data) ?? [],
          isProtected: !!commentDataByTag["protected"]?.length,
          throws: commentDataByTag["throws"]?.map((c) => c.data) ?? [],
          requires: commentDataByTag["requires"]?.map((c) => c.data) ?? [],
          isStatic: !!commentDataByTag["static"]?.length
        };
        arrToPushTo.push(func);
        break;
      case "property":
        let prop = {
          name: sName,
          access,
          description: commentDataByTag["description"]?.[0]?.data ?? "",
          remarks: commentDataByTag["remark"]?.map((c) => c.data) ?? [],
          examples: commentDataByTag["example"]?.map((c) => c.data) ?? [],
          params,
          returns: {
            type: commentDataByTag["returns"]?.[0]?.data.type ?? sRetType,
            description: commentDataByTag["returns"]?.[0]?.data.description ?? ""
          },
          deprecation: {
            status: !!commentDataByTag["deprecated"]?.length,
            message: commentDataByTag["deprecated"]?.[0]?.data ?? ""
          },
          isDefaultMember: defaultMember === sName,
          devNotes: commentDataByTag["devNote"]?.map((c) => c.data) ?? [],
          todos: commentDataByTag["todo"]?.map((c) => c.data) ?? [],
          isProtected: !!commentDataByTag["protected"]?.length,
          throws: commentDataByTag["throws"]?.map((c) => c.data) ?? [],
          requires: commentDataByTag["requires"]?.map((c) => c.data) ?? [],
          isStatic: !!commentDataByTag["static"]?.length
        };
        properties.push(prop);
        break;
      case "event":
        events.push({
          name: sName,
          description: commentDataByTag["returns"]?.[0]?.data.description ?? "",
          remarks: commentDataByTag["remark"]?.map((c) => c.data) ?? [],
          examples: commentDataByTag["example"]?.map((c) => c.data) ?? [],
          params,
          devNotes: commentDataByTag["devNote"]?.map((c) => c.data) ?? [],
          todos: commentDataByTag["todo"]?.map((c) => c.data) ?? []
        });
        break;
    }
  });
  let mod = {
    name: moduleName,
    fileName,
    description: moduleDocsByTag["description"]?.[0]?.data ?? "",
    remarks: moduleDocsByTag["remark"]?.map((c) => c.data) ?? [],
    examples: moduleDocsByTag["example"]?.map((c) => c.data) ?? [],
    devNotes: moduleDocsByTag["devNote"]?.map((c) => c.data) ?? [],
    todos: moduleTODOs,
    requires: moduleDocsByTag["requires"]?.map((c) => c.data) ?? [],
    methods,
    properties
  };
  if (isClass) {
    return {
      ...mod,
      constructors,
      events,
      implements: _implements
    };
  } else {
    return mod;
  }
}
function main() {
  let files = fs.readdirSync(__dirname + "/../../src");
  files = files.filter(
    (f) => fs.lstatSync(__dirname + "/../../src/" + f).isFile()
  );
  let docs = files.map((file) => {
    return parseModuleOrClass(
      fs.readFileSync(__dirname + "/../../src/" + file, "utf8"),
      file
    );
  });
  fs.writeFileSync(
    __dirname + "/../../docs.json",
    JSON.stringify(docs, null, 2),
    "utf8"
  );
}
main();
//# sourceMappingURL=main.js.map
