'Simple interface to create XML:
  ''Example usage
  'Dim sXML as string
  'With stdXML.Create("A")
  '  .pushChild(stdXML.Create("B")).setAttributes("A",1,"B","Hi").setValue(1)
  '  .pushChild(stdXML.Create("B")).setAttributes("A",2).setValue("And stuff")
  'End With
  ''==>
  ''<A>
  ''  <B A=1, B="Hi">1</B>
  ''  <B A=2>And stuff</B>
  ''</A>

  ''Real life example (sharepoint where condition)
  'Dim sXML as string
  'With stdXML.Create("Query")
  '  With .pushChild(stdXML.Create("Where"))
  '    With .pushChild(stdXML.Create("Eq"))
  '      .pushChild(stdXML.Create("FieldRef")).setAttributes("Name","LinkFilename")
  '      .pushChild(stdXML.Create("Value")).setAttributes("Type","Text").setValue("test.txt")
  '    End With
  '  End With
  '
  '  'Serialize
  '  sXML = .ToString()
  'End With
  ''==>
  ''<Query>
  ''  <Where>
  ''    <Eq>
  ''      <FieldRef Name="LinkFilename"/>
  ''      <Value Type="Text">test.txt</Value>
  ''    </Eq>
  ''  </Where>
  ''</Query>

'Embedding other data structures:
  'Embed JSON
  'Dim sXML as string
  'With stdXML.Create("Data")
  '  With .pushChild(stdXML.Create("JSON"))
  '    With stdJSON.cast( .pushChildEx(stdJSON.Create()))
  '      .setValue("A",1)
  '      With .createArrayElement("B")
  '        .pushChild(1)
  '        .pushChild(2)
  '        .pushChild(3)
  '      End With
  '    End With
  '  End With
  '  sXML = .ToString()
  'End With
  ''==>
  ''<Data>
  ''  <JSON>
  ''    {
  ''      "A":1,
  ''      "B":[1,2,3]
  ''    }
  ''  </JSON>
  ''</Data>

  ''Embedding custom ISerialisable objects
  'With stdXML.Create("Data")
  '  .pushChildEx(myObject)
  '  Debug.Print .ToString
  'End With
  ''<Data>
  ''  Whatever bro
  ''</Data>

Public Function Create(ByVal sTagName as string) as stdXML

End Function
Public Function CreateTextNode(ByVal sValue as string) as stdXML
  'for use in cases like <A><B/>hello<C/>world</A>
End Function
Public Function CreateFromString(ByVal sXML as string) as stdXML

End Function
Public Function CreateFromFile(ByVal sXML as string) as stdXML

End Function
Public Function ToString() as string

End Function


Public Function pushChild(obj as ISerialisable) as stdXML

End Function
Public Function pushChildEx(obj as ISerialisable) as object

End Function