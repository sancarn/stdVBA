VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'stdTable myTable
' |- stdArray rows
'    |- stdTableRow row
'       |- stdDictionary data
'
' |
'
'
'

class stdTable
  Implements ISerializable
  Implements IIterable
  
  Private WithEvents pRows as stdArray
  Private WithEvents pColumns as stdArray
  
  
  Public Event OnRowAdd(row as stdTableRow, rowid as long)
  Public Event OnRowRemove(row as stdTableRow, rowid as long)
  Public Event OnRowChange(row as stdTableRow, rowid as long, changeType as string, optional key as string, optional value as variant)
  
  
  Function Create(???) as stdTable
  
  Function FromListObject(lo as ListObject) as stdTable
  Function FromArray(vArr as variant, Optional metadata as object = nothing) as stdTable
  Function FromExcelRange(rng as range, optional metadata as object = nothing) as stdTable
  Function FromText(data as string, optional metadata as object = nothing) as stdTable
  Function FromFile(data as string, optional metadata as object = nothing) as stdTable
  Function FromArrayOfDicts(arr as stdArray, optional metadata as object) as stdTable
  Function FromJSON(arr as stdArray, optional metadata as object) as stdTable
  
  Function ToListObject(lo as ListObject) as ListObject
  Function ToArray(Optional metadata as object = nothing) as stdArray<stdDictionary>
  Function ToExcelRange(rng as range) as range
  Function ToText(optional metadata as object = nothing) as String
  Function ToFile(optional metadata as object = nothing) as stdFile
  Function ToJSON() as string
  
  Function LinkTo(lo as ListObject) as boolean 'Attempts to link virtual table to list object events
  
  Function ISerializable_ToString(): ISerializable_ToString=ToJSON: End Function
  
  Function IIterable_ForEach(stdCallback): IIterable_ForEach = pRows.ForEach(stdCallback): End Function
  Function IIterable_Map(stdCallback)    : IIterable_Map     = pRows.Map(stdCallback)    : End Function
  '...
  
  
  Property Headers as stdArray
  
  Property ColumnValidation as stdDictionary
  
  
  Property Rows as stdArray
  Property Columns as stdDictionary
  
  Function pRows_OnAdded(item as variant, index as long)
  Function pColumns_OnAdded(item as variant, index as long)
  
end class


class stdArray
  Public Event OnAdded(item as variant, index as long)
  Public Event OnRemoved(item as variant, index as long)
  Public Event OnChange()
  
end class