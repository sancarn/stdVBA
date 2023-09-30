
```
| A great form                           | X |
|---------------------------------------------
| Name                                       |
| |-----------------|                        |
| |                 |                        |
| |-----------------|                        |
|                                            |
| | Needed? | How Many? |                    |
| |---------|-----------|                    |
| |   ▢     |          |                    |
| |   ▢     |          |                    |
| |   ▢     |          |                    |
| |   ▢     |          |                    |
| |   ▢     |          |                    |
| |   ▢     |          |                    |
| |   ▢     |          |                    |
| |   ▢     |          |                    |
| |---------------------|                    |
|---------------------------------------------
```


```vb
With uiForm.Create("A great form")
    With .addChild uiBox.Create(layout:=uiRows, padding:=5)
        .addChild uiInputBox.Create("Name")
        With .AddChild(uiTable.Create(Array(uiTableHeader.Create("Needed?", uiCheckbox), uiTableHeader.Create("How Many?", uiEdit))))
            For each row in myData
                .AddChild uiTableRow.Create(row("isRequired"), row("count"))
            next
        End With
    End With
End With
```