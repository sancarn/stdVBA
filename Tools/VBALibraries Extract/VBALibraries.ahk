;Written in AHK 1.1.30.00


#Include libs\JSON.ahk
#SingleInstance,Force
ControlGet, lb1, hwnd,, ListBox1, ahk_class wndclass_desked_gsk ahk_exe EXCEL.EXE
ControlGet, lb2, hwnd,, ListBox2, ahk_class wndclass_desked_gsk ahk_exe EXCEL.EXE
ControlGet, txt1, hwnd,, RichEdit20A1, ahk_class wndclass_desked_gsk ahk_exe EXCEL.EXE
ControlGet, cb1, hwnd,, ComboBox1, ahk_class wndclass_desked_gsk ahk_exe EXCEL.EXE
ControlGet, cb1_list, list,,ComboBox1, ahk_class wndclass_desked_gsk ahk_exe EXCEL.EXE

libCount := 0
Loop, Parse, cb1_list, `n
  libCount += 1

i:=0, j:=0
errLib:=0,errA:=0, errB:=0

;Main return object:
libraries := {type:"Libraries",zChildren:[]}

;Start from first library...
k:=1 ;Skip 0 (<AllLibraries>)

While k < libCount {
  ;Increment library item to target
  k +=1
  
  ;Get library
  ;Note: you need to show and hide the dropdown for the description field to update in this case...
  Control,Choose,%k%,, ahk_id %cb1%
  Control, ShowDropDown ,,, ahk_id %cb1%
  Control, HideDropDown ,,, ahk_id %cb1%
  library := parseLibrary()
  
  
  ErrA:=0,i:=0
  While errA==0 {
    ;Increment listbox item to target
    i += 1
    
    ;Try to get control:
    Control,Choose,%i%,, ahk_id %lb1%
    errA := ErrorLevel
    
    ;Get Parent:
    parent := parseParent()
    
    j:=0,errB:=0
    While errB==0 {
      ;Increment listbox item to target
      j+=1
      
      ;Try to get control
      Control,Choose,%j%,, ahk_id %lb2%
      errB:=ErrorLevel
      
      ;Parse member and append to parent
      parent.members.push(parseMember())
    }
    
    ;Append parent and members to array:
    library.zChildren.push(parent)
  }
  
  ;Push library to libraries object
  libraries.zChildren.push(library)
}

;Stringify and dump to VBALibraries.json
sJSON := JSON.stringify(libraries)
FileDelete,%A_ScriptDir%\VBALibraries.json
FileAppend,%sJSON%, %A_ScriptDir%\VBALibraries.json

msgbox Exporting VBALibraries has complete
return



parseParent(){
  haystack := getDefinition()
  needle   := "iO)(\<globals\>)|([^ ]+) ([^ ]+)\s+Member of (.+)"
  RegexMatch(haystack,needle, oMatch)
  isGlobals := StrLen(oMatch.Value(1))>0
  
  if isGlobals {
    return {type: "global",name: "<globals>", parent:"", members:[],typeDescription:haystack}
  } else {
    type  := oMatch.value(2)
    name  := oMatch.value(3)
    parent:= oMatch.value(4)
    return {type: type, name:name, parent:parent, members:[],typeDescription:haystack}
  }
  
}
parseMember(){
  haystack := getDefinition()
  needle    = O)(?:(?<MemberType>[^ ]+) )?(?<MemberName>[^ (]+)\(?(?<Params>.*?)?\)?(?: As (?<RetType>[^ ]+))?\s*(?:(?<ReadOnly>read-only))?\s*(?<Default>Default )?[mM]ember of (?<Parent>.+)
  RegexMatch(haystack,needle, oMatch)
  isDefault  := StrLen(oMatch.Default)>0
  isReadOnly := StrLen(oMatch.ReadOnly)>0
  MemberType := oMatch.MemberType
  MemberName := oMatch.MemberName
  ParamString:= oMatch.Params
  RetType    := oMatch.RetType ? oMatch.RetType : "VOID"
  Parent     := oMatch.Parent
  
  return {MemberType:MemberType,MemberName:MemberName,ParamString:ParamString,RetType:RetType,Parent:Parent,isDefault:isDefault,isReadOnly:isReadOnly,typeDescription:haystack}
}
parseLibrary(){
  haystack := getDefinition()
  needle = O)(?<LibType>[^ ]+) (?<LibName>[^ ]+)(?:\s+(?<LibPath>.+)\s+(?<LibDescription>.+))?
  RegexMatch(haystack,needle, oMatch)
  LibType        := oMatch.LibType       
  LibName        := oMatch.LibName       
  LibPath        := oMatch.LibPath       
  LibDescription := oMatch.LibDescription
  
  return {LibType:LibType,LibName:LibName,LibPath:LibPath,LibDescription:LibDescription,zChildren:[]}
}


getDefinition(){
  global
  ControlGetText, text,, ahk_id %txt1%
  return text
}