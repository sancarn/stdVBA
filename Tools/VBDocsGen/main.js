/**
* TODO: Figure out how to handle conditional compilation
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
* TODO: Add implemented interfaces
* TODO: Add public types?
* TODO: Add public constants

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

function groupBy(list, keyGetter) {
    const map = {};
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

/**
 * Transforms a comment into a comment store object
 * @param {string} comment - The comment to parse
 * @returns {CommentStore} - The comment store object
 * @typedef {(CommentRecord[])} CommentStore
 * @typedef {({type: ICommentType, data: ICommentData})} CommentRecord
 * @typedef {null|IParamData|IReturnData|IExampleData} ICommentData
 * @typedef {{name: string, type: string, description: string}} IParamData //Note: Optional, ByVal/ByRef, ParamArray, DefaultValue obtained from code
 * @typedef {{type: string, description: string}} IReturnData
 * @typedef {{example: string}} IExampleData
 * @typedef {"constructor"|"protected"|"deprecated"|"param"|"returns"|"example"|"description"} ICommentType
 */
function parseComment(comment) {
    //inject @description into 1st line of comment
    comment = comment.replace(/^'/, "'@description ");

    const groupByRx = /'@(?<type>\w+)(?<content>.*\s+(?:'[^@][^\n]*\s+)*)/g;
    const regexTags = {
        description: /'@description\s+(?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
        param: /'@param\s+(?<name>\w+)\s*(?:as\s+(?<type>[^-]+))?(?:\s*-\s*(?<description>.+\s*(?:'[^@][^\n]*\n?)*))?/i, //regex needs work
        returns: /'@returns\s+(?<type>[^-\r\n]+)?(?:\s*-\s+(?<description>.+\s*(?:'[^@][^\n]*\n?)*))?/i,
        example: /'@example\s+(?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
        remark: /'@remark\s+(?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
        constructor: /@constructor/i,
        protected: /@protected/i,
        deprecated: /@deprecated (?<description>.+\s*(?:'[^@][^\n]*\n?)*)/i,
        defaultMember: /@defaultMember/i
    };

    //Parse comment into comment store
    const commentStore = [];
    const matches = [...comment.matchAll(groupByRx)];
    for(match of matches) {
        let type = match.groups.type;
        if(!!regexTags[type]){
            let data = regexTags[type].exec(match[0])?.groups;
            //If comment not valid ignore
            if(!!data){
                if(!!data?.description) data.description = data.description.trim().replace(/^'/gm, "");
                commentStore.push({type, data});
            }
        } else {
            commentStore.push({type, data: match.groups.content.trim().replace(/^'/gm, "")});
        }
    }

    return commentStore;
}

/**
 * Parse params string into params array
 * @param {string} params 
 * @param {CommentGroups} commentGroups 
 * @typedef {"param": {[key: string]: IParamData},[key: ICommentType]: ICommentData[]} CommentGroups
 */
function parseParams(params, commentGroups) {
    const paramRegex = /(?<optional>optional\s+)?(?:(?<referenceType>byval|byref)\s+)?(?:(?<paramarray>paramarray)\s+)?(?<name>\w+)(?<isArray>\(\))?(?:\s+as\s+(?<type>[^, )]+))?(?:\s*=\s*(?<defaultValue>[^,\)]+))?/gi;
    const paramMatches = [...params.matchAll(paramRegex)];
    const paramStore = [];
    for(paramMatch of paramMatches) {
        const comment = commentGroups?.param?.[paramMatch.groups.name.toLowerCase()];
        //Parse param type, prioritise comment type, then code type, then default to Variant
        let paramType;
        if(!!comment?.data.type){
            paramType = comment?.data.type;
        } else if(!!paramMatch.groups.type) {
            if(paramMatch.groups.isArray){
                paramType = `Array<${paramMatch.groups.type}>`;
            } else {
                paramType = paramMatch.groups.type;
            }
        } else {
            paramType = "Variant";
        }
        paramStore.push({
            name: paramMatch.groups.name.trim(),
            type: paramType.trim(),
            description: comment?.data.description.trim() ?? "",
            referenceType: paramMatch.groups?.referenceType ?? "ByRef",
            paramArray: !!paramMatch.groups.paramArray,
            optional: !!paramMatch.groups.optional,
            defaultValue: (!!paramMatch.groups.optional ? (paramMatch.groups.defaultValue ?? "Unspecified") : null)
        });
    }

    return paramStore;
}

//Find all files in ../../src directory
let fs = require("fs");
let files = fs.readdirSync(__dirname + "/../../src");
files = files.filter(f => fs.lstatSync(__dirname + "/../../src/" + f).isFile());
let modules = [];

//Scan files for public methods and properties, dump to docs array
files.forEach(file => {
    let content = fs.readFileSync(__dirname + "/../../src/" + file, "utf8");
    //remove all conditional compilation
    content = content.replace(/#if.+then\s+(.+)(.|\s)+?#end if/gi,"$1");

    //Initialise module
    const moduleNameFinder = /Attribute VB_Name = "([^"]+)"/i;
    const myModule = {
        name: moduleNameFinder.exec(content)[1],
        fileName: file,
        methods: []
    };

    const docsFinder = /(?<comments>(?:\'.*\r?\n)*)Public (?:(?<type>Function|Sub|Event|Property) ?(?<access>Get|Let|Set)?) (?<name>\w+)\((?<params>[^)]*)\)(?: as (?<retType>\w+))?/gmi;
    const matches = Array.from(content.matchAll(docsFinder));
    matches.forEach(match => {
        if(match.processed) return;

        if(match.groups.name.toLowerCase() == "findfirst"){
            1==1;
        }

        const commentData = parseComment(match.groups.comments);
        
        //Convert commentData to object with type as key, and param name as subkey of param
        const commentGroups = groupBy(commentData, c => c.type);
        if(!!commentGroups.param){
            commentGroups.param = groupBy(commentGroups.param, c => c.data.name.toLowerCase());
            Object.keys(commentGroups.param).forEach(name => commentGroups.param[name] = commentGroups.param[name][0]);
        }
        
        const params = parseParams(match.groups.params, commentGroups);

        //obtain access from scanning matches for Get/Let/Set. If Get only, access is ReadOnly. If Let/Set only, access is WriteOnly. If both, access is ReadWrite.
        let access = null;
        if(match.groups.type.toLowerCase() === "property"){
            access = match.groups.access.toLowerCase() === "get" ? "ReadOnly" : "WriteOnly";
            let others = [];
            switch(access){
                case "ReadOnly":
                    others = matches.filter(m => m.groups.name === match.groups.name && ["let","set"].includes(m.groups.access.toLowerCase()))
                    break;
                case "WriteOnly":
                    others = matches.filter(m => m.groups.name === match.groups.name && ["get"].includes(m.groups.access.toLowerCase()))
                    break;
            }
            //If there are other matches, this is a ReadWrite property and we need to mark the others as processed
            if(others.length > 0){
                access = "ReadWrite";
                others.forEach(m => {
                    m.processed = true;
                });
            }
        }

        //obtain return type from comment if present, otherwise from code
        let returnType = match.groups.retType;
        if(!!commentGroups.returns) returnType = commentGroups.returns[0].data.type;
        if(match.groups.type.toLowerCase() == "sub") returnType = "Void";

        //Add method to module
        myMethod = {
            name: match.groups.name,
            type: match.groups.type,
            constructor: !!commentGroups.constructor,
            isDefaultMember: !!commentGroups.defaultMember,
            protected: !!commentGroups.protected,
            access,
            description: commentGroups?.description?.[0]?.data?.description ?? "",
            params: params,
            returns: {
                type: returnType,
                description: commentGroups?.returns?.[0]?.data?.description ?? ""
            },
            deprecation: {
                status: !!commentGroups.deprecated,
                message: commentGroups?.deprecated?.[0]?.data?.description ?? ""
            },
            examples: commentGroups?.example?.map(c => c.data.description) ?? [],
            remarks: commentGroups?.remark?.map(c => c.data.description) ?? []
        };
        console.log(myMethod);
        myModule.methods.push(myMethod);
    })

    modules.push(myModule);
});

//Dump docs array to docs.json
fs.writeFileSync(__dirname + "/../../docs.json", JSON.stringify(modules, null, 2), "utf8");


