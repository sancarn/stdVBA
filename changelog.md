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
    * `NOTE` - any disclaimers etc.
* `<Short description>` - Short description of the change / fix / feature
* `<BODY>` to be provided if required

Before `08/07/2021` a change log was not kept. We have  retrospectively gone back and populated a change log to `04/05/2020`. Any changes before this date will be missing from the change log, however they will still be identifiable from decent commit comments in `github history`. If interested use `git log --all`

# Change log since `2020-05-04`


* ...
* 2020-05-04 stdAcc         WIP         - `stdAcc` added + first commit.
* 2020-05-15 stdDate        BREAKING    - `stdDate::Create2` renamed to `stdDate::CreateFromUnits()`
* 2020-05-15 stdDate        BREAKING    - `stdDate::Parse` renamed `stdDate::CreateFromParse()` as with `stdVBA` convention
* 2020-05-15 stdDate        FEATURE     - `stdDate::CreateFromParse` can now parse `dd.mm.yyyy` as well as `dd/mm/yyyy` efficiently.
* 2020-05-15 stdDate        FEATURE     - `stdDate::CreateFromParse` calculates `y`, `m`, `d`, `h`, `mn`, `s` individually and then uses ``stdDate::CreateFromUnits()` to generate date.
* 2020-05-31 STD            BREAKING    - Remove reference to `STD`. Long time deprecated.
* 2020-06-10 stdDate        FIX         -  Correction in `stdDate`
* 2020-06-27 stdArray       FIX         -  `stdArray#Init` now takes params `byval` - `byref` was causing issues where it'd change the values in other classes - not good!
* 2020-06-27 stdArray       FIX         -  `stdArray#Clone` needed to pass `pLength` over to the new array, this wasn't being done correctly. Fixed.
* 2020-06-27 stdArray       FIX         -  `stdArray#Reduce` now uses `isMissing` on metadata and the optional initialValue is given as 0.
* 2020-06-27 stdCallaback   DEPRECATED  - `stdCallaback#Create()` deprecated
* 2020-06-27 stdCallaback   FEATURE     - `stdCallaback#CreateFromModule()` added
* 2020-06-27 stdCallaback   FEATURE     - `stdCallaback#CreateFromObjectMethod()` added
* 2020-06-27 stdCallaback   FEATURE     - `stdCallaback#CreateFromObjectProperty()` added
* 2020-07-28 stdCallaback   FIX         - Spelling error
* 2020-08-01 stdRegex       FIX         - Fix for crash in `stdRegex`
* 2020-08-01 stdArray       FEATURE     - Added `stdArray#remove` 
* 2020-08-01 stdArray       FIX         - `byval` fix for `stdArray#concat`
* 2020-08-17 stdLambda      FEATURE     - Added `stdLambda` (moved to `src` from `wip`).
* 2020-08-17 stdLambda      FEATURE     - Added `ThisWorkbook` and `Application` to `stdLambda` as keyword constants.
* 2020-08-19 stdLambda      FIX         - Use VBE7 instead of `msvbvm60` for `rtcCallByname` in `stdLambda`. `msvbvm60` isn't always available.
* 2020-08-24 stdArray       FEATURE     - Added events to `stdArray`
* 2020-08-30 stdLambda      FEATURE     - `stdLambda` now uses a VM evaluation approach. Code is first tokenised, parsed, compiled to bytecode. Byte code evaluated when executed.
* 2020-09-01 stdCallback    FIX         - `stdCallback::Create()` fix - call type passed as wrong param to init.
* 2020-09-01 stdArray       FIX         - `stdArray#Unique` - Fix vKeys(i) to vKeys.item(i)
* 2020-09-10 stdICallable   FEATURE     - `stdICallable` added and `stdLambda` and `stdCallback` implement it.
* 2020-09-10 stdLambda      FEATURE     - `stdLambda#run` is now default method of stdLambda
* 2020-09-10 stdArray       FEATURE     - `stdArray#sort` added
* 2020-09-10 stdArray       FEATURE     - `stdArray#item` is now default method of array
* 2020-09-10 stdArray       FIX         - `stdArray#arr()` will return an initialised array of zero length when length is 0.
* 2020-09-14 stdArray       FIX         - `stdArray#pLength` not reducing on shift
* 2020-09-14 stdArray       FIX         - `stdArray#Unshift` incorrect index used.
* 2020-09-14 stdArray       FIX         - Fix bug with missing SortStruct in stdArray
* 2020-09-14 stdArray       FIX         - Fix potential crash - avoid using copy memory in Property Get Arr()
* 2020-09-16 stdLambda      FEATURE     - `stdLambda#bind()` and `stdCallback#bind()` added.
* 2020-09-16 stdICallable   BREAKING    - Call convention of `stdICallable#RunEx()` to use `ByVal array`.
* 2020-09-16 stdLambda      FIX         - Typo `isObject(stdLambda) = "Empty"` instead of `isObject(stdLambda)`
* 2020-09-16 stdLambda      FIX         - Typo `stdCallback.CreateEvaluator("1") is stdLambda` instead of `TypeOf stdCallback.CreateEvaluator("1") is stdICallable`
* 2020-09-16 stdLambda      FIX         - `stdLambda` used to call functions in reverse parameter order. This has been fixed.
* 2020-09-25 stdLambda      FIX         - Fix `and` behaves like `or` in `stdLambda`
* 2020-10-03 stdLambda      FEATURE     - Added `switch()` and `any()` functions to `stdLambda`.
* 2020-10-05 stdEnumerator  FEATURE     - Added `stdEnumerator`.
* 2020-10-12 stdClipboard   FEATURE     - Added `stdClipboard`.
* 2020-10-12 stdLambda      WIP         - Moving towards `stdLambda` Mac compatibility.
* 2020-11-11 stdWindow      WIP         - `stdWindow` first commit.
* 2020-11-11 stdShell       WIP         - `stdShell` first commit.
* 2020-11-11 stdArray       FEARURE     - Added `stdArray#Min()` and `stdArray#Max()` functions.
* 2020-11-13 stdWindow      WIP         - Large number of `stdWindow` additions
* 2020-11-13 stdArray       BREAKING    - Switch from using `stdArray` return value to `Collection`.
* 2020-11-13 stdRegex       BREAKING    - Removal of `stdRegex::Create2()` pending use.
* 2020-11-13 stdRegex       FEATURE     - Added `stdRegex#ListArr()` which is an easy method of creating 2d arrays of data from regex matches
* 2020-11-15 stdRegex       FIX         - Fix bug with `stdRegex` - needed to get type information in order to call friend method.
* 2020-11-15 stdArray       BREAKING    - Removal of callback metadata. Use `stdICallable#Bind()` instead.
* 2020-11-15 _Various       FIX         - Many bugs fixed after the introduction of unit testing.
* 2020-12-08 stdLambda      FIX         - Better keyword matching for `stdLambda`.
* 2020-12-08 stdArray       FEATURE     - Added an optional starting value parameter to `stdArray#Min()` and `stdArray#Max()` functions.
* 2020-12-15 stdPerformance FEARURE     - Added `stdPerformance` class.
* 2021-02-11 stdEnumerator  FEATURE     - Added `stdEnumerator#asCollection()` and `stdEnumerator#asArray()`.
* 2021-02-12 stdWebSocket   FEATURE     - Added `stdWebSocket`.
* 2021-02-12 stdClipboard   FIX         - `OpenClipboard` now uses `OpenClipboardTimeout`. Opening clipboard can timeout, and is detectable.
* 2021-03-01 stdLambda      FEATURE     - Added `stdLambda#BindGlobal()`
* 2021-03-01 stdLambda      FEATURE     - Added `Dictionary.Key` syntax to `stdLambda`
* 2021-03-01 stdICallable   DEV         - `stdICallable#SendMessage()` added. Not advised that people depend on this function as it is technically internal. Offers latebinding for stdICallable objects.
* 2021-03-01 stdAcc         BREAKING    - Renamed `stdAcc::FromWindow`, `stdAcc::FromIUnknown`, ... to `stdAcc::CreateFromWindow`, `stdAcc::CreateFromIUnknown`, ... to be inline with `stdVBA` standards.
* 2021-03-01 stdAcc         BREAKING    - `stdAcc#FindFirst` and `stdAcc#FindAll` now use `stdICallable` instead of query parameters.
* 2021-03-01 stdAcc         FEATURE     - `EAccStates`, `EAccRoles` and `EAccFindResult` are injected into `stdICallable`s which support the `bindGlobal()` method (currently `stdLambda` alone)
* 2021-03-01 stdAcc         BREAKING    - `stdAcc::CreateFromExcel()` renamed to `stdAcc::CreateFromApplication()` as this function now also works in Word.
* 2021-03-01 stdAcc         BREAKING    - Changed code format of `stdAcc#Text()` to a JSON-like format.
* 2021-03-01 stdAcc         FEATURE     - Added `stdAcc#PrintChildTexts()` and `stdAcc#PrintDescTexts()` which are useful when debugging.
* 2021-03-01 stdAcc         FIX         - Proxy parent now returns `stdAcc` instead of `IAccessible`
* 2021-03-01 stdAcc         FIX         - `Role` and `State` changed to use new system
* 2021-03-01 stdAcc         FIX         - Safer handling of `WindowFromAccessibleObject`
* 2021-03-03 stdAcc         FIX         - Compile error fixes.
* 2021-03-11 stdProcess     FEATURE     - Added `stdProcess`.
* 2021-03-20 stdWindow      FEATURE     - Added `stdWindow`.
* 2021-03-27 stdLambda      FEATURE     - Added a performance cache to `stdLambda`, which increases the speed of result evaluation in certain cases.
* 2021-03-27 stdLambda      FIX         - Small bug fixes to `evaluateFunc()`
* 2021-04-09 stdLambda      FEATURE     - Added `null`, `nothing`, `empty` and `missing` to `stdLambda`.
* 2021-04-09 stdEnumerator  FEATURE     - Added `stdEnumerator::CreateFromCallable()`
* 2021-04-09 stdEnumerator  FEATURE     - Added `stdEnumerator::CreateFromArray()`
* 2021-04-09 stdEnumerator  BREAKING    - Added callback parameter to `stdEnumerator#unique`. BREAKING fixed in patch on 2021-04-10
* 2021-04-09 stdEnumerator  FEATURE     - Added the `like` operator to `stdLambda`
* 2021-04-10 stdEnumerator  FEATURE     - Made callback of `stdEnumerator#unique` optional
* 2021-04-10 stdEnumerator  FEATURE     - Made callback of `stdEnumerator#sort` optional
* 2021-04-10 stdEnumerator  BREAKING    - `init` renamed to `protInit`. Unlikely to affect anyone.
* 2021-04-10 stdEnumerator  BREAKING    - `withIndex` optional parameters removed from `stdEnumerator` and instead callbacks are always passed the index. Any usage of `stdCallback` will now need to implement a 2nd and 3rd parameter for the key and index.
* 2021-04-10 stdEnumerator  FIX         - Fix typo in `stdEnumerator#NextItem()` where `CallableVerbose` returned data to the wrong array on callback execute.
* 2021-04-10 stdEnumerator  FIX         - Fixes to stdEnumeratorTests.bas to ensure all tests succeed
* 2021-04-11 stdProcess     FIX         - `stdProcess`'s `Time` functions would crash if `pQueryInfoHandle=0`. Add a check and exit property.
* 2021-04-11 _UnitTests     FIX         - Fixes to Main test file
* 2021-04-21 stdLambda      FIX         - Move `stdLambda`'s `Like` operator from `iType.oMisc` to  `iType.oComparison`
* 2021-05-05 stdCOM         FEATURE     - Added stdCOM
* 2021-05-18 stdLambda      FIX         - Ensure `stdLambda.oFuncExt` is always defined.
* 2021-05-21 stdDictionary  WIP         - Initial work to `stdDictionary`
* 2021-05-21 stdTable       WIP         - Initial work to `stdTable`
* 2021-06-18 stdLambda      BREAKING    - Fixed Unintuitive right-to-left behavior of `stdLambda`. This is theoretically breaking, however unlikely to affect anyone negatively. Ultimately `8/2/2` will now return `2` (as it is running the equivalent of `(8/2)/2`) instead of `8` (as it used to run the equivalent of `8/(2/2)`). I.E. The change makes Math work as it does in VBA and most other programming languages.
* 2021-06-27 stdArray       FIX         - `stdArray#arr()` should use `CopyVariant` instead of `=`
* 2021-06-27 stdEnumerator  FEATURE     - `stdEnumerator::CreateEmpty()` added.
* 2021-06-27 stdEnumerator  FIX         - `stdEnumerator#protInit()` now works for 0-length enumerators.
* 2021-07-02 stdEnumerator  FIX         - Fixed `stdEnumerator#AsArray()` works even if `stdEnumerator` is empty.
* 2021-07-02 stdWindow      BREAKING    - Fixed `stdWindow#X()`, `stdWindow#Y()`,`stdWindow#Width()`,`stdWindow#Height()` to relate to RectClient instead of RectWindow. 
* 2021-07-02 stdEnumerator  FEATURE     - Added optional parameter to `stdEnumerator#FindFirst()` which will return if the item is not found.
* 2021-07-06 stdEnumerator  FEATURE     - Added `stdEnumerator::CreateFromListObject()`
* 2021-07-06 stdCOMDispatch WIP         - Started work on `IDispatch` wrapper using `stdCOM` 
* 2021-07-07 _UnitTests     FIX         - Fixed bug in testing environment. Ensured that `Test.Range` existed in mainBuilder.
* 2021-07-07 stdLambda      BREAKING    - `#` is no longer valid inside `stdLambda` expression. Use `.` (for method OR property access), `.$` (for property specific access) or `.#` (for method specific access). I.E. If you have code like `obj#method` you should change this to `obj.method` as `.`. In some rare cases you may have to use `.#` instead.
* 2021-07-07 stdLambda      FEATURE     - Added `pEquation` property to `stdLambda` - useful while debugging.
* 2021-07-08 stdEnumerator  FIX         - Fixed an issue where `stdEnumerator#Sort()` through an error on empty arrays
* 2021-07-08 stdEnumerator  FIX         - Fixed an issue where `stdEnumerator#AsArray()` wouldn't return an array of the correct type when used with anything other than `VbVariant` as argument.
* 2021-07-09 stdCallback    FIX         - Fixed an issue where `CriticalRaise` would occur in `stdCallback`, ending runtime, where it actually successfully ran.
* 2021-07-09 stdClipboard   FIX         - Fixed typo in `GetPictureFromClipboard()` from `if OpenClipboardTimeOut()>1 then` to `if OpenClipboardTimeOut() then`
* 2021-07-09 stdWindow      BREAKING    - Reverting `2021-07-02 BREAKING` change. Use of `WindowRect` should be default, however `ClientRect` should also be allowed. See next line for new feature.
* 2021-07-09 stdWindow      FEATURE     - Added optional Rect type parameter to x,y,width and height. Use `wnd.x(RectTypeClient) = ...` to modify with respect to the client rect.
* 2021-07-10 stdLambda      FIX         - Remove TODO statement from `stdLambda` evaluation loop. 
* 2021-07-10 stdLambda      FIX         - Check for `Application` and `ThisWorkbook` existence in `stdLambda`. This brings `Word` and `VB6` compatibility.  
* 2021-07-10 stdProcess     BREAKING    - All protected methods in `stdProcess` are now declared as `Friend` instead of `Public`.
* 2021-07-10 stdProcess     FIX         - Removed `stdProcess#moduleID` as it was always returning `0`. Need to look into how to get `moduleID`s in a class based setting.
* 2021-07-10 stdProcess     NOTE        - Added documentation note to all Time functions of stdProcess e.g. `stdProcess#TimeCreated()`, indicating that this function currently always returns time in UTC timezone.
* 2021-07-10 stdProcess     BREAKING    - `stdProcess::getProcessImageName` set to `Private`. This function should never have been public. Replace with `stdProcess.Create(...).path`
* 2021-07-10 stdEnumerator  FIX         - Fixed issue with `stdEnumerator::CreateFromListObject()` - compile error due to lack of test. Test added nowand 100% working.
* 2021-07-18 stdEnumerator  FEATURE     - Added `stdEnumerator#AsArray2D()`.
* 2021-08-05 stdPerformance FEATURE     - Added optional parameter to stdPerformance which acts as a divisor for the final time. I.E. `totalTime/nCount`. Useful where you also loop internally over something to get a time per operation.