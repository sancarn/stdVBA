# stdPerformance

## Scoped application optimisation

~~~vb
With stdPerformance.CreateOptimiser()
  'Bulk work here.
End With
~~~

Excel temporarily disables ScreenUpdating and EnableEvents and selects manual Calculation, then restores the captured values. Word changes and restores ScreenUpdating. Access maps ScreenUpdating to Application.Echo and does not access Excel-only EnableEvents or Calculation properties.

## Access repainting and nesting

~~~vb
With stdPerformance.CreateOptimiser()
  'Access repainting is disabled.
  With stdPerformance.CreateOptimiser(ScreenUpdating:=True)
    'Repainting is temporarily enabled.
  End With
  'The outer request is active again.
End With
'Repainting is enabled.
~~~

Access exposes Echo as a method with no readable current-state property. The optimiser therefore assumes repainting should be enabled when the outermost scope finishes. If calling code already disabled Echo, declare the desired exit state:

~~~vb
Application.Echo False
With stdPerformance.CreateOptimiser(AccessEchoOnExit:=False)
  'Bulk work.
End With
'Echo remains off; the caller owns re-enabling it.
Application.Echo True
~~~

Nested optimiser scopes track the latest active request, including when retained optimiser objects are released out of order. The outermost scope's AccessEchoOnExit determines the final state. Changes made by unrelated direct Echo calls while scopes are active cannot be tracked.

## Other Access options

Echo is the general application-wide repainting switch. The default optimiser deliberately limits itself to repainting:

- Form.Painting controls a particular form, with separate handling required for subforms. It is useful for targeted layout changes.
- DoCmd.SetWarnings suppresses messages and confirmation prompts. It changes behavior and requires its own explicit cleanup; it is not an event or calculation switch.
- Name AutoCorrect options affect saved database metadata and dependency tracking. Changing them can require closing objects or reopening the database, so they are database-design decisions rather than scoped runtime toggles.

Access has no application-wide EnableEvents or Excel-style Calculation property.

References: [Echo](https://learn.microsoft.com/en-us/office/vba/api/access.application.echo), [Painting](https://learn.microsoft.com/en-us/office/vba/api/access.form.painting), [SetWarnings](https://learn.microsoft.com/en-us/office/vba/api/access.docmd.setwarnings), [Name AutoCorrect](https://support.microsoft.com/en-us/access/set-name-autocorrect-options).

## Tests

Run stdPerformanceAccessTests.testAll inside Access. Run stdPerformanceHostTests.testAll inside Excel or Word; it is also included in MainTestAll. Existing Excel measurement tests remain in stdPerformanceTests.

