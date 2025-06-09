/**
     'Some comment
     '@param MethodPointer as LongPtr - Pointer to the method to call
     '@param iRetType as VbVarType - The return type of the method
     '@param vParamTypes as Array<Variant<VbVarType>> - The types of the parameters
     #If VBA7 Then
       Public Function CreateFromPointer(ByVal MethodPointer As LongPtr, Optional ByVal iRetType As VbVarType = 0, Optional vParamTypes = Null) As stdCallback
     #Else
       Public Function CreateFromPointer(ByVal MethodPointer As Long, Optional ByVal iRetType As VbVarType = 0, Optional vParamTypes = Null) As stdCallback
     #End If
* TODO: Add public enums
* TODO: Add public types?
* TODO: Add public constants
* TODO: Add @throws    e.g. @throws 1, "ERROR: Only stdReg keys have subkeys"
* TODO: Add @requires  e.g. @requires stdLambda
*/

/**
 * Take input stubs like:
 *
 *   'Creates an `stdAcc` object from an `X` and `Y` point location on the screen.
 *   '@constructor
 *   '@protected
 *   '@deprecated
 *   '@param x as Long - X Coordinate
 *   '@param y as Long - Y Coordinate
 *   '@returns stdAcc - Object at the specified point
 *   '@example `acc.CreateFromPoint(100, 200).DoDefaultAction`
 *   '@example ```
 *   '  acc.CreateFromPoint(0, 0).FindFirst(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))
 *   '```
 *   Public Function CreateFromPoint(ByVal x As Long, ByVal y As Long) As stdAcc
 *
 *   'Search the IAccessible tree for elements which match a certain criteria. Return the first element found.
 *   '@param ByVal query as stdICallable<(stdAcc,depth)=>EAccFindResult> - Callback returning
 *   '  EAccFindResult options:
 *   '    EAccFindResult.NoMatchFound/0/False             - Not found, countinue walking
 *   '    EAccFindResult.MatchFound/1/-1/True             - Found, return this element
 *   '    EAccFindResult.NoMatchCancelSearch/2            - Not found, cancel search
 *   '    EAccFindResult.NoMatchSkipDescendents/3,else    - Not found, don't search descendents
 *   '    EAccFindResult.MatchFoundSearchDescendents/4    - Same as EAccFindResult.MatchFound
 *   '@param {EAccFindType=1} - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
 *   ' To understand the difference between BFS and DFS take this tree:
 *   '        A
 *   '       / \
 *   '      B   C
 *   '     /   / \
 *   '    D   E   F
 *   ' A BFS will walk this tree in the following order: A, B, C, D, E, F
 *   ' A DFS will walk this tree in a different order:   A, C, F, E, B, D
 *   '@examples
 *   ' ```
 *   ' 'Find where name is "hello" and class is "world":
 *   ' el.FindFirst(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))
 *   ' 'Find first element named "hello" at depth > 4:
 *   ' el.FindFirst(stdLambda.Create("$1.name = ""hello"" AND $2 > 4"))
 *   ' ```
 *   Public Function FindFirst(ByVal query As stdICallable, optional byval searchType as EAccFindType=EAccFindType.DepthFirst) As stdAcc
 *
 * And dump these to docs array, with structure:
 * [
 *   {
 *     name: string //name in VBattributes code
 *     methods: [
 *       {
 *         name: string //name in code
 *         type: "sub"|"function"|"property"
 *         params: [
 *           {
 *             name: string //name in code
 *             type: string //type in code or type in comment if present
 *             description: string //description in comment,
 *             optional: boolean //true if param is optional
 *             defaultValue: string //default value if param is optional
 *             paramArray: boolean //true if param is paramarray, false otherwise
 *           },
 *           ...
 *         ],
 *         returns: {
 *           type: string //return type in code or type in comment if present
 *           description: string //description in comment
 *         },
 *         description: string //description in comment
 *         access: "ReadOnly"|"WriteOnly"|"ReadWrite" //access in code
 *         protected: boolean //true if protected as per comment
 *         deprecated: boolean //true if deprecated as per comment
 *         constructor: boolean //true if constructor as per comment
 *         examples: string[] //examples in comment as markdown
 *       },
 *       ...
 *     ]
 *   }
 * ]
 */

function log(
  message: string,
  type: "info" | "warn" | "error" | "success" = "info"
) {
  switch (type) {
    case "info":
      console.log(`\x1b[36mℹ️  Info: ${message}\x1b[0m`);
      break;
    case "warn":
      console.log(`\x1b[33;1m⚠️  Warn: ${message}\x1b[0m`);
      break;
    case "error":
      console.log(`\x1b[31;1m❌  Error: ${message}\x1b[0m`);
      break;
    case "success":
      console.log(`\x1b[32;1m✅  Success: ${message}\x1b[0m`);
      break;
  }
}

/**
 * Groups an array of objects by a key getter
 * @param list - The array to group
 * @param keyGetter - The key getter to use to group the array
 * @returns - The grouped array
 */
function groupBy<T>(
  list: T[],
  keyGetter: (item: T) => string
): { [key: string]: T[] } {
  const map = Object.create(null) as { [key: string]: T[] };
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

type IDocs = (IModule | IClass)[];
type IModule = {
  name: string;
  fileName: string;
  description: string;
  remarks: string[];
  examples: string[];
  methods: IMethod[];
  properties: IProperty[];
  devNotes: string[];
  todos: string[];
  requires: string[];
};
type IClass = IModule & {
  implements: string[];
  constructors: IConstructor[];
  events: IEvent[];
};
type IMethod = {
  name: string;
  description: string;
  remarks: string[];
  examples: string[];
  params: IParam[];
  returns: IReturn;
  devNotes: string[];
  todos: string[];
  throws: IThrows[];
  requires: string[];

  isStatic: boolean;
  isProtected: boolean;
  isDefaultMember: boolean;
  deprecation: {
    status: boolean;
    message: string;
  };
};

type IProperty = IMethod & {
  access: "ReadOnly" | "WriteOnly" | "ReadWrite";
};
type IConstructor = IMethod;
type IEvent = {
  name: string;
  description: string;
  remarks: string[];
  examples: string[];
  params: IParam[];
  devNotes: string[];
  todos: string[];
};
type IParam = {
  name: string;
  type: string;
  description: string;
  optional: boolean;
  defaultValue: string | null;
  paramArray: boolean;
};
type IReturn = {
  type: string;
  description: string;
};
type IThrows = {
  errNumber: number;
  errText: string;
};

type ITagTypes =
  | "constructor"
  | "protected"
  | "deprecated"
  | "param"
  | "returns"
  | "example"
  | "description"
  | "remark"
  | "devNote"
  | "TODO"
  | "throws"
  | "requires"
  | "static";

type ITagLine = {
  tag: ITagTypes;
  content: string;
};

type ICommentStore = ICommentRecord[];
type ICommentRecord =
  | IDataConstructor
  | IDataProtected
  | IDataDeprecated
  | IDataParam
  | IDataReturn
  | IDataExample
  | IDataDescription
  | IDataRemark
  | IDataDevNote
  | IDataTODO
  | IDataThrows
  | IDataRequires
  | IDataStatic;

type IDataConstructor = {
  tag: "constructor";
};
type IDataProtected = {
  tag: "protected";
};
type IDataDeprecated = {
  tag: "deprecated";
  data: "" | string; //E.G. "Use `stdLambda.Create()` instead."
};
type IDataParam = {
  tag: "param";
  data: {
    name: string; //from comment
    type: string; //from comment or param data
    description: string; //from comment
    referenceType?: string; //from param data
    paramArray?: boolean; //from param data
    optional?: boolean; //from param data
    defaultValue?: string | null; //from param data
  };
};
type IDataReturn = {
  tag: "returns";
  data: {
    type: string; //from comment or param data
    description: string;
  };
};
type IDataExample = {
  tag: "example";
  data: string;
};
type IDataDescription = {
  tag: "description";
  data: string;
};
type IDataRemark = {
  tag: "remark";
  data: string;
};
type IDataDevNote = {
  tag: "devNote";
  data: string;
};
type IDataTODO = {
  tag: "TODO";
  data: string;
};
type IDataThrows = {
  tag: "throws";
  data: {
    errNumber: number;
    errText: string;
  };
};
type IDataRequires = {
  tag: "requires";
  data: string;
};
type IDataStatic = {
  tag: "static";
};

//Assertions to ensure all tags declared in ITagTypes are implemented
type TagFromRecord = ICommentRecord extends { tag: infer T } ? T : never;
type TagsAreEqual = [ITagTypes] extends [TagFromRecord]
  ? [TagFromRecord] extends [ITagTypes]
    ? true
    : false
  : false;
type _AssertTagMatch<T extends true> = T;

// ❌ Error here if not all tags in ITagTypes are implemented in ICommentRecord
type __triggerTagMismatchError = _AssertTagMatch<TagsAreEqual>;

/**
 * Parses a comment block and extracts lines starting with tagged annotations.
 *
 * A tagged line must begin with `'@tagName` followed by its content.
 * Lines without a leading `'@` are grouped as part of the previous tag's content.
 *
 * Example input:
 * ```
 * '@test hello
 * '@test world
 * '@this is
 * 'fine
 * '@thing
 * ```
 * Produces:
 * ```ts
 * [
 *   { tag: "test", content: "hello" },
 *   { tag: "test", content: "world" },
 *   { tag: "this", content: "is\r\n'fine" },
 *   { tag: "thing", content: "" }
 * ]
 * ```
 * @param comment - The full comment string to parse.
 * @returns An array of tag-content pairs extracted from the comment.
 */
function parseToTagLines(comment: string): ITagLine[] {
  const groupByRx = /'@(?<type>\w+)(?<content>.*\s+(?:'[^@][^\n]*\s+)*)?/g;
  const matches = [...comment.matchAll(groupByRx)];
  return matches.map((match) => ({
    tag: match.groups?.type as ITagTypes,
    content: match.groups?.content?.replace(/^'/gm, "") ?? "",
  }));
}

/**
 * Transforms a comment into a comment store object
 * @param comment - The comment to parse
 * @returns - The comment store object
 */
function parseComment(comment: string): ICommentStore {
  //If undefined or empty, return empty array
  if (!comment) return [];

  //inject @description into 1st line of comment for easier parsing
  comment = comment.replace(/^'/g, "'@description ");
  const tagLines = parseToTagLines(comment);

  //Extracts and groups comments under their flag/tag type e.g. "@example hello\r\n'world"
  const regexTags = {
    description: /(?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
    param:
      /(?<name>\w+)\s*(?:as\s+(?<type>[^-]+))?(?:\s*-\s*(?<description>.+\s*(?:'[^@][^\n]*\n?)*))?/i, //regex needs work
    returns:
      /(?<type>[^-\r\n]+)?(?:\s*-\s+(?<description>.+\s*(?:'[^@][^\n]*\n?)*))?/i,
    example: /(?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
    remark: /(?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
    deprecated: /(?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
    devNote: /(?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
    TODO: /(?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
    constructor: /(?:constructor)?/g, //overwrites native constructor
    throws: /(?<errNumber>\d+)\s*,\s*(?<errText>.+)/i,
    requires: /(?<description>.+)/i,
  };

  //Parse comment into comment store
  const commentStore: ICommentStore = [];
  for (let tagLine of tagLines) {
    let tag = tagLine.tag;
    let groups;
    if (!!regexTags[tag]) {
      groups = regexTags[tag].exec(tagLine.content)?.groups;
    }

    //If comment not valid ignore
    switch (tag) {
      case "description":
        if (!!groups?.description)
          commentStore.push({ tag, data: groups.description });
        break;
      case "param":
        if (!!groups?.name)
          commentStore.push({
            tag,
            data: {
              name: groups.name,
              type: groups?.type,
              description: groups?.description,
            },
          });
        break;
      case "returns":
        commentStore.push({
          tag,
          data: {
            type: groups?.type,
            description: groups?.description,
          },
        });
        break;
      case "example":
        commentStore.push({ tag, data: groups?.description });
        break;
      case "remark":
        commentStore.push({ tag, data: groups?.description });
        break;
      case "devNote":
        commentStore.push({ tag, data: groups?.description });
        break;
      case "constructor":
        commentStore.push({ tag });
        break;
      case "protected":
        commentStore.push({ tag });
        break;
      case "deprecated":
        commentStore.push({ tag, data: groups?.description });
        break;
      case "TODO":
        commentStore.push({ tag, data: groups?.description });
        break;
      case "throws":
        commentStore.push({
          tag,
          data: {
            errNumber: Number(groups?.errNumber),
            errText: groups?.errText,
          },
        });
        break;
      case "requires":
        commentStore.push({ tag, data: groups?.description });
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

/**
 * Given a VBA param string and a param data object initialised from the comment, obtains additional information from the param data and injects it into the param data object.
 * @param params - The param string to parse.
 * @param dataParams - The params data objects to populate.
 * @returns - The populated param data object.
 */
function parseParams(params: string, dataParams: IDataParam[]): IDataParam[] {
  //Group params by name
  const paramData = groupBy<IDataParam>(dataParams, (d) =>
    d.data.name.toLowerCase()
  );

  const paramRegex =
    /(?<optional>optional\s+)?(?:(?<referenceType>byval|byref)\s+)?(?:(?<paramarray>paramarray)\s+)?(?<name>\w+)(?<isArray>\(\))?(?:\s+as\s+(?<type>[^, )]+))?(?:\s*=\s*(?<defaultValue>[^,\)]+))?/gi;
  const paramMatches = [...params.matchAll(paramRegex)];

  const paramStore: IDataParam[] = [];
  for (let paramMatch of paramMatches) {
    //Get param name from Function/Sub declaration
    let name = paramMatch.groups?.name;
    if (!name) continue;

    //Get param data from comment
    if (!paramData[name.toLowerCase()]?.length) continue;
    const commentData = paramData[name.toLowerCase()][0].data;

    /**
     * Parse param type, prioritise comment type, then code type, then default to Variant
     * description is always the comment description
     * referenceType is from the function/sub declaration, or ByRef if unspecified
     * paramArray is from the function/sub declaration, or false if unspecified
     * optional is from the function/sub declaration, or false if unspecified
     * If param is optional, and no default value is provided, set default value to null
     */
    let paramType: string =
      commentData.type ?? paramMatch.groups?.type ?? "Variant";
    let paramDescription: string = commentData.description ?? "";
    let paramReferenceType: string =
      paramMatch.groups?.referenceType ?? "ByRef";
    let paramIsParamArray: boolean = !!paramMatch.groups?.paramArray;
    let paramIsArray: boolean = !!paramMatch.groups?.isArray;
    let paramIsOptional: boolean = !!paramMatch.groups?.optional;
    let paramDefaultValue: string | null = !!paramMatch.groups?.optional
      ? paramMatch.groups.defaultValue ?? null
      : null;

    //Special cases
    //If param is an array, and no type is provided in the commentary, assume it's an array of `paramType`
    if (paramIsArray && !!commentData.type) paramType = `Array<${paramType}>`;

    //Add param to store
    paramStore.push({
      tag: "param",
      data: {
        name: name.trim(),
        type: paramType.trim(),
        description: paramDescription.trim(),
        referenceType: paramReferenceType,
        paramArray: paramIsParamArray,
        optional: paramIsOptional,
        defaultValue: paramDefaultValue,
      },
    });
  }

  return paramStore;
}

function parseModuleOrClass(
  content: string,
  fileName: string
): IModule | IClass {
  let isClass = /^VERSION 1.0 CLASS/.test(content);

  //remove all conditional compilation from module
  let regexConditionalCompilation = /#if.+then\s+((.|\s)+?)#end if/gi;
  while (regexConditionalCompilation.test(content)) {
    content = content.replace(regexConditionalCompilation, "$1");
  }

  //Initialise module
  const moduleNameFinder = /Attribute VB_Name = "(?<name>[^"]+)"/i;
  const moduleName =
    moduleNameFinder.exec(content)?.groups?.name ?? fileName.split(".")[0];
  log(`Parsing module "${moduleName}"`);
  const moduleDocsFinder = /'@module.*\r?\n('.*\r?\n)*/i;
  const moduleDocsString = moduleDocsFinder.exec(content)?.groups?.[0];
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

  //Find documentation. This usually looks like a comment block, followed by a method/property/event declaration.
  const memberDocsFinder =
    /(?<comments>(?:\'.*\r?\n)*)(?<!' *)Public (?:(?<type>Function|Sub|Event|Property) ?(?<access>Get|Let|Set)?) (?<name>\w+)\((?<params>[^)]*)\)(?: as (?<retType>\w+))?/gim;
  const memberMatches = Array.from(content.matchAll(memberDocsFinder));

  //populate members
  let constructors: IConstructor[] = [];
  let events: IEvent[] = [];
  let properties: IProperty[] = [];
  let methods: IMethod[] = [];
  let membersByName = groupBy(memberMatches, (m: any) =>
    m.groups?.name.toLowerCase()
  ); //used for property access
  let memberAlreadyProcessed: { [key: string]: boolean } = {};
  memberMatches.forEach((match) => {
    let sComment = match.groups?.comments;
    let sType: "function" | "sub" | "event" | "property" =
      match.groups?.type.toLowerCase() as
        | "function"
        | "sub"
        | "event"
        | "property";
    let sName = match.groups?.name;
    let sParams = match.groups?.params ?? "";
    let sRetType = match.groups?.retType ?? "Void";

    //Mark member as processed to ensure it doesn't get processed again, especially important for properties as each `get`, `let` and `set` are declared as different rows.
    if (memberAlreadyProcessed[sName.toLowerCase()]) return;
    memberAlreadyProcessed[sName.toLowerCase()] = true;

    //Determine property access
    let access: "ReadOnly" | "WriteOnly" | "ReadWrite";
    if (sType === "property") {
      let accessTypes = membersByName[sName.toLowerCase()].map((match) =>
        match.groups?.access.toLowerCase()
      );
      let readAccess = accessTypes.includes("get");
      let writeAccess =
        accessTypes.includes("let") || accessTypes.includes("set");
      if (readAccess && writeAccess) {
        access = "ReadWrite";
      } else if (readAccess) {
        access = "ReadOnly";
      } else if (writeAccess) {
        access = "WriteOnly";
      }
    }

    let commentData: ICommentStore;
    if (sComment.length > 0) {
      commentData = parseComment(sComment);
    }
    let commentDataByTag = groupBy(commentData, (c) => c.tag);

    let params: IParam[] = parseParams(
      sParams,
      commentDataByTag["param"] as IDataParam[]
    ).map((param: IDataParam) => {
      return {
        name: param.data.name,
        type: param.data.type,
        description: param.data?.description ?? "",
        optional: param.data?.optional ?? false,
        defaultValue: param.data?.defaultValue ?? null,
        paramArray: param.data?.paramArray ?? false,
      };
    });

    switch (sType) {
      case "sub":
      case "function":
        let arrToPushTo: IMethod[] | IConstructor[] = !!commentDataByTag[
          "constructor"
        ]?.length
          ? constructors
          : methods;

        let func: IConstructor | IMethod = {
          name: sName,
          description:
            (commentDataByTag["description"]?.[0] as IDataDescription)?.data ??
            "",
          remarks:
            commentDataByTag["remark"]?.map((c: IDataRemark) => c.data) ?? [],
          examples:
            commentDataByTag["example"]?.map((c: IDataExample) => c.data) ?? [],
          params,
          returns:
            sType === "sub"
              ? null
              : {
                  type:
                    (commentDataByTag["returns"]?.[0] as IDataReturn)?.data
                      .type ?? sRetType,
                  description:
                    (commentDataByTag["returns"]?.[0] as IDataReturn)?.data
                      .description ?? "",
                },
          deprecation: {
            status: !!commentDataByTag["deprecated"]?.length,
            message:
              (commentDataByTag["deprecated"]?.[0] as IDataDeprecated)?.data ??
              "",
          },
          isDefaultMember: defaultMember === sName,
          devNotes:
            commentDataByTag["devNote"]?.map((c: IDataDevNote) => c.data) ?? [],
          todos: commentDataByTag["todo"]?.map((c: IDataTODO) => c.data) ?? [],
          isProtected: !!commentDataByTag["protected"]?.length,
          throws:
            commentDataByTag["throws"]?.map((c: IDataThrows) => c.data) ?? [],
          requires:
            commentDataByTag["requires"]?.map((c: IDataRequires) => c.data) ??
            [],
          isStatic: !!commentDataByTag["static"]?.length,
        };
        arrToPushTo.push(func);
        break;
      case "property":
        let prop: IProperty = {
          name: sName,
          access,
          description:
            (commentDataByTag["description"]?.[0] as IDataDescription)?.data ??
            "",
          remarks:
            commentDataByTag["remark"]?.map((c: IDataRemark) => c.data) ?? [],
          examples:
            commentDataByTag["example"]?.map((c: IDataExample) => c.data) ?? [],
          params,
          returns: {
            type:
              (commentDataByTag["returns"]?.[0] as IDataReturn)?.data.type ??
              sRetType,
            description:
              (commentDataByTag["returns"]?.[0] as IDataReturn)?.data
                .description ?? "",
          },
          deprecation: {
            status: !!commentDataByTag["deprecated"]?.length,
            message:
              (commentDataByTag["deprecated"]?.[0] as IDataDeprecated)?.data ??
              "",
          },
          isDefaultMember: defaultMember === sName,
          devNotes:
            commentDataByTag["devNote"]?.map((c: IDataDevNote) => c.data) ?? [],
          todos: commentDataByTag["todo"]?.map((c: IDataTODO) => c.data) ?? [],
          isProtected: !!commentDataByTag["protected"]?.length,
          throws:
            commentDataByTag["throws"]?.map((c: IDataThrows) => c.data) ?? [],
          requires:
            commentDataByTag["requires"]?.map((c: IDataRequires) => c.data) ??
            [],
          isStatic: !!commentDataByTag["static"]?.length,
        };
        properties.push(prop);
        break;
      case "event":
        events.push({
          name: sName,
          description:
            (commentDataByTag["returns"]?.[0] as IDataReturn)?.data
              .description ?? "",
          remarks:
            commentDataByTag["remark"]?.map((c: IDataRemark) => c.data) ?? [],
          examples:
            commentDataByTag["example"]?.map((c: IDataExample) => c.data) ?? [],
          params,
          devNotes:
            commentDataByTag["devNote"]?.map((c: IDataDevNote) => c.data) ?? [],
          todos: commentDataByTag["todo"]?.map((c: IDataTODO) => c.data) ?? [],
        });
        break;
    }
  });

  //Build base module
  let mod: IModule = {
    name: moduleName,
    fileName,
    methods,
    properties,
    description:
      (moduleDocsByTag["description"]?.[0] as IDataDescription)?.data ?? "",
    remarks: moduleDocsByTag["remark"]?.map((c: IDataRemark) => c.data) ?? [],
    examples:
      moduleDocsByTag["example"]?.map((c: IDataExample) => c.data) ?? [],
    devNotes:
      moduleDocsByTag["devNote"]?.map((c: IDataDevNote) => c.data) ?? [],
    todos: moduleTODOs,
    requires:
      moduleDocsByTag["requires"]?.map((c: IDataRequires) => c.data) ?? [],
  };

  //If it's a class then add additional members
  if (isClass) {
    return {
      ...mod,
      constructors,
      events,
      implements: _implements,
    };
  } else {
    return mod;
  }
}

import { throws } from "assert";
import * as fs from "fs";
function main() {
  //Find all files in ../../src directory
  let files = fs.readdirSync(__dirname + "/../../src");
  files = files.filter((f) =>
    fs.lstatSync(__dirname + "/../../src/" + f).isFile()
  );

  //Scan files for public methods and properties, dump to docs array
  let docs: IDocs = files.map((file) => {
    return parseModuleOrClass(
      fs.readFileSync(__dirname + "/../../src/" + file, "utf8"),
      file
    );
  });

  //Dump docs array to docs.json
  fs.writeFileSync(
    __dirname + "/../../docs.json",
    JSON.stringify(docs, null, 2),
    "utf8"
  );
}

main();
