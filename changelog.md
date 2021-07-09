# Changelog

Change log shall be a bullet point list. All changes will be of the format:

```
* yyyy-mm-dd <TAG> - <Short description>

<BODY>
```

Where:

* `<TAG>` is to be either
    * `FIX` - For bug fixes
    * `FEATURE` - For features added
    * `BREAKING` when a breaking change occurs which will break software using this feature/bug.
    * `WIP` if work in progress still
    * `DEV` a feature which should only be utilised by `stdVBA` developers.
    * `DEPRECATED` when a feature still supported but is deprecated.
* `<Short description>` - Short description of the change / fix / feature
* `<BODY>` to be provided if required

Before `08/07/2021` a change log was not kept. We have  retrospectively gone back and populated a change log to `04/05/2020`. Any changes before this date will be missing from the change log, however they will still be identifiable from decent commit comments in `github history`. If interested use `git log --all`

# Change log since `2020-05-04`


* ...
* 2020-05-04 WIP         - `stdAcc` added + first commit.
* 2020-05-15 BREAKING    - `stdDate::Create2` renamed to `stdDate::CreateFromUnits()`
* 2020-05-15 BREAKING    - `stdDate::Parse` renamed `stdDate::CreateFromParse()` as with `stdVBA` convention
* 2020-05-15 FEATURE     - `stdDate::CreateFromParse` can now parse `dd.mm.yyyy` as well as `dd/mm/yyyy` efficiently.
* 2020-05-15 FEATURE     - `stdDate::CreateFromParse` calculates `y`, `m`, `d`, `h`, `mn`, `s` individually and then uses ``stdDate::CreateFromUnits()` to generate date.
* 2020-05-31 BREAKING    - Remove reference to `STD`. Long time deprecated.
* 2020-06-10 FIX         -  Correction in `stdDate`
* 2020-06-27 FIX         -  `stdArray#Init` now takes params `byval` - `byref` was causing issues where it'd change the values in other classes - not good!
* 2020-06-27 FIX         -  `stdArray#Clone` needed to pass `pLength` over to the new array, this wasn't being done correctly. Fixed.
* 2020-06-27 FIX         -  `stdArrayReduce` now uses `isMissing` on metadata and the optional initialValue is given as 0.
* 2020-06-27 DEPRECATED  - `stdCallaback#Create()` deprecated
* 2020-06-27 FEATURE     - `stdCallaback#CreateFromModule()` added
* 2020-06-27 FEATURE     - `stdCallaback#CreateFromObjectMethod()` added
* 2020-06-27 FEATURE     - `stdCallaback#CreateFromObjectProperty()` added
* 2020-06-27 FEATURE     - `stdCallabackMixin` added (contains a lambda-like workaround which uses formulae)
* 2020-07-28 FIX         - Spelling error
* 2020-08-01 FIX         - Fix for crash in `stdRegex`
* 2020-08-01 FEATURE     - Added `stdArray#remove` 
* 2020-08-01 FIX         - `byval` fix for `stdArray#concat`
* 2020-08-17 FEATURE     - Added `stdLambda` (moved to `src` from `wip`).
* 2020-08-17 FEATURE     - Added `ThisWorkbook` and `Application` to `stdLambda` as keyword constants.
* 2020-08-19 FIX         - Use VBE7 instead of `msvbvm60` for `rtcCallByname` in `stdLambda`. `msvbvm60` isn't always available.
* 2020-08-24 FEATURE     - Added events to `stdArray`
* 2020-08-30 FEATURE     - `stdLambda` now uses a VM evaluation approach. Code is first tokenised, parsed, compiled to bytecode. Byte code evaluated when executed.
* 2020-09-01 FIX         - `stdCallback::Create()` fix - call type passed as wrong param to init.
* 2020-09-01 FIX         - `stdArray#Unique` - Fix vKeys(i) to vKeys.item(i)
* 2020-09-10 FEATURE     - `stdICallable` added and `stdLambda` and `stdCallback` implement it.
* 2020-09-10 FEATURE     - `stdLambda#run` is now default method of stdLambda
* 2020-09-10 FEATURE     - `stdArray#sort` added
* 2020-09-10 FEATURE     - `stdArray#item` is now default method of array
* 2020-09-10 FIX         - `stdArray#arr()` will return an initialised array of zero length when length is 0.
* 2020-09-14 FIX         - `stdArray#pLength` not reducing on shift
* 2020-09-14 FIX         - `stdArray#Unshift` incorrect index used.
* 2020-09-14 FIX         - Fix bug with missing SortStruct in stdArray
* 2020-09-14 FIX         - Fix potential crash - avoid using copy memory in Property Get Arr()
* 2020-09-16 FEATURE     - `stdLambda#bind()` and `stdCallback#bind()` added.
* 2020-09-16 BREAKING    - Call convention of `stdICallable#RunEx()` to use `ByVal array`.
* 2020-09-16 FIX         - Typo `isObject(stdLambda) = "Empty"` instead of `isObject(stdLambda)`
* 2020-09-16 FIX         - Typo `stdCallback.CreateEvaluator("1") is stdLambda` instead of `TypeOf stdCallback.CreateEvaluator("1") is stdICallable`
* 2020-09-16 FIX         - `stdLambda` used to call functions in reverse parameter order. This has been fixed.
* 2020-09-25 FIX         - Fix `and` behaves like `or` in `stdLambda`
* 2020-10-03 FEATURE     - Added `switch()` and `any()` functions to `stdLambda`.
* 2020-10-05 FEATURE     - Added `stdEnumerator`.
* 2020-10-12 FEATURE     - Added `stdClipboard`.
* 2020-10-12 WIP         - Moving towards `stdLambda` Mac compatibility.
* 2020-11-11 WIP         - `stdWindow` first commit.
* 2020-11-11 WIP         - `stdShell` first commit.
* 2020-11-11 FEARURE     - Added `stdArray#Min()` and `stdArray#Max()` functions.
* 2020-11-13 WIP         - Large number of `stdWindow` additions
* 2020-11-13 BREAKING    - Switch from using `stdArray` return value to `Collection`.
* 2020-11-13 BREAKING    - Removal of `stdRegex::Create2()` pending use.
* 2020-11-13 FEATURE     - Added `stdRegex#ListArr()` which is an easy method of creating 2d arrays of data from regex matches
* 2020-11-15 FIX         - Fix bug with `stdRegex` - needed to get type information in order to call friend method.
* 2020-11-15 BREAKING    - Removal of callback metadata. Use `stdICallable#Bind()` instead.
* 2020-11-15 FIX         - Many bugs fixed after the introduction of unit testing.
* 2020-12-08 FIX         - Better keyword matching for `stdLambda`.
* 2020-12-08 FEATURE     - Added an optional starting value parameter to `stdArray#Min()` and `stdArray#Max()` functions.
* 2020-12-15 FEARURE     - Added `stdPerformance` class.
* 2021-02-11 FEATURE     - Added `stdEnumerator#asCollection()` and `stdEnumerator#asArray()`.
* 2021-02-12 FEATURE     - Added `stdWebSocket`.
* 2021-02-12 FIX         - `OpenClipboard` now uses `OpenClipboardTimeout`. Opening clipboard can timeout, and is detectable.
* 2021-03-01 FEATURE     - Added `stdLambda#BindGlobal()`
* 2021-03-01 FEATURE     - Added `Dictionary.Key` syntax to `stdLambda`
* 2021-03-01 DEV         - `stdICallable#SendMessage()` added. Not advised that people depend on this function as it is technically internal. Offers latebinding for stdICallable objects.
* 2021-03-01 BREAKING    - Renamed `stdAcc::FromWindow`, `stdAcc::FromIUnknown`, ... to `stdAcc::CreateFromWindow`, `stdAcc::CreateFromIUnknown`, ... to be inline with `stdVBA` standards.
* 2021-03-01 BREAKING    - `stdAcc#FindFirst` and `stdAcc#FindAll` now use `stdICallable` instead of query parameters.
* 2021-03-01 FEATURE     - `EAccStates`, `EAccRoles` and `EAccFindResult` are injected into `stdICallable`s which support the `bindGlobal()` method (currently `stdLambda` alone)
* 2021-03-01 BREAKING    - `stdAcc::CreateFromExcel()` renamed to `stdAcc::CreateFromApplication()` as this function now also works in Word.
* 2021-03-01 BREAKING    - Changed code format of `stdAcc#Text()` to a JSON-like format.
* 2021-03-01 FEATURE     - Added `stdAcc#PrintChildTexts()` and `stdAcc#PrintDescTexts()` which are useful when debugging.
* 2021-03-01 FIX         - Proxy parent now returns `stdAcc` instead of `IAccessible`
* 2021-03-01 FIX         - `Role` and `State` changed to use new system
* 2021-03-01 FIX         - Safer handling of `WindowFromAccessibleObject`
* 2021-03-03 FIX         - Compile error fixes.
* 2021-03-11 FEATURE     - Added `stdProcess`.
* 2021-03-20 FEATURE     - Added `stdWindow`.
* 2021-03-27 FEATURE     - Added a performance cache to `stdLambda`, which increases the speed of result evaluation in certain cases.
* 2021-03-27 FIX         - Small bug fixes to `evaluateFunc()`
* 2021-04-09 FEATURE     - Added `null`, `nothing`, `empty` and `missing` to `stdLambda`.
* 2021-04-09 FEATURE     - Added `stdEnumerator::CreateFromCallable()`
* 2021-04-09 FEATURE     - Added `stdEnumerator::CreateFromArray()`
* 2021-04-09 BREAKING    - Added callback parameter to `stdEnumerator#unique`. BREAKING fixed in patch on 2021-04-10
* 2021-04-09 FEATURE     - Added the `like` operator to `stdLambda`
* 2021-04-10 FEATURE     - Made callback of `stdEnumerator#unique` optional
* 2021-04-10 FEATURE     - Made callback of `stdEnumerator#sort` optional
* 2021-04-10 BREAKING    - `init` renamed to `protInit`. Unlikely to affect anyone.
* 2021-04-10 BREAKING    - `withIndex` optional parameters removed from `stdEnumerator` and instead callbacks are always passed the index.
* 2021-04-10 FIX         - Fix typo in `stdEnumerator#NextItem()` where `CallableVerbose` returned data to the wrong array on callback execute.
* 2021-04-10 FIX         - Fixes to stdEnumeratorTests.bas to ensure all tests succeed
* 2021-04-11 FIX         - `stdProcess`'s `Time` functions would crash if `pQueryInfoHandle=0`. Add a check and exit property.
* 2021-04-11 FIX         - Fixes to Main test file
* 2021-04-21 FIX         - Move `stdLambda`'s `Like` operator from `iType.oMisc` to  `iType.oComparison`
* 2021-05-05 FEATURE     - Added stdCOM
* 2021-05-18 FIX         - Ensure `stdLambda.oFuncExt` is always defined.
* 2021-05-21 WIP         - Initial work to `stdTable` and `stdDictionary`
* 2021-06-18 BREAKING    - Fixed Unintuitive right-to-left behavior of `stdLambda`. This is theoretically breaking, however unlikely to affect anyone negatively.
* 2021-06-27 FIX         - `stdArray#arr()` should use `CopyVariant` instead of `=`
* 2021-06-27 FEATURE     - `stdEnumerator::CreateEmpty()` added.
* 2021-06-27 FIX         - `stdEnumerator#protInit()` now works for 0-length enumerators.
* 2021-07-02 FIX         - Fixed `stdEnumerator#AsArray()` works even if `stdEnumerator` is empty.
* 2021-07-02 BREAKING    - Fixed `stdWindow#X()`, `stdWindow#Y()`,`stdWindow#Width()`,`stdWindow#Height()` to relate to RectClient instead of RectWindow. 
* 2021-07-02 FEATURE     - Added optional parameter to `stdEnumerator#FindFirst()` which will return if the item is not found.
* 2021-07-06 FEATURE     - Added `stdEnumerator::CreateFromListObject()`
* 2021-07-06 WIP         - Started work on `IDispatch` wrapper using `stdCOM` 
* 2021-07-07 FIX         - Fixed bug in testing environment. Ensured that `Test.Range` existed in mainBuilder.
* 2021-07-07 BREAKING    - `#` is no longer valid inside `stdLambda` expression. Use `.`, `.$` or `.#` instead.
* 2021-07-07 FEATURE     - Added `pEquation` property to `stdLambda` - useful while debugging.
* 2021-07-08 FIX         - Fixed an issue where `stdEnumerator#Sort()` through an error on empty arrays
* 2021-07-08 FIX         - Fixed an issue where `stdEnumerator#AsArray()` wouldn't return an array of the correct type when used with anything other than `VbVariant` as argument.
* 2021-07-09 FIX         - Fixed an issue where `CriticalRaise` would occur in `stdCallback`, ending runtime, where it actually successfully ran.
* 2021-07-09 FIX         - Fixed typo in `GetPictureFromClipboard()` from `if OpenClipboardTimeOut()>1 then` to `if OpenClipboardTimeOut() then`