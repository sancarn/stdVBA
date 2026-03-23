# `stdWebView`

`stdWebView` embeds the Microsoft Edge **WebView2** control in a Win32 parent window. It is aimed at **Excel VBA UserForms**: host the browser inside an `MSForms.Frame` (or any HWND you obtain) to show HTML, open sites, and run JavaScript with optional async callbacks.

**Platform:** Windows only. Requires the [WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) and `WebView2Loader.dll` (same loader used by the WebView2 SDK) available to the host process.

**Implementation note:** Environment and controller callbacks are implemented with in-process vtables and executable thunks (similar patterns appear elsewhere in stdVBA). See the class header in `src/stdWebView.cls` for attribution to prior WebView2/VBA work.

## Spec

### Constructors

#### `CreateFromHwnd(ByVal hwnd As LongPtr, Optional ByVal OnReady As stdICallable = Nothing) As stdWebView`

Creates a `stdWebView` bound to an existing window handle. The WebView fills the **client area** of that window.

`OnReady`, if provided, is invoked once the controller and `CoreWebView2` are ready. The callback is called via `stdICallable.RunEx` with a single-element array: the `stdWebView` instance (`Array(Me)`).

```vb
Dim wv As stdWebView
Set wv = stdWebView.CreateFromHwnd(SomeHwnd, stdLambda.Create("$1.Navigate ""https://example.com"""))
```

Construction is **synchronous from the caller’s perspective**: the factory polls `DoEvents` until initialization finishes or times out (raises `WebView2 initialization failed or timed out`).

#### `CreateFromFrame(ByVal frm As Object, Optional ByVal OnReady As stdICallable = Nothing) As stdWebView`

Same as `CreateFromHwnd`, but resolves the HWND from an **`MSForms.Frame`**. Raises if `frm` is not a `Frame`.

```vb
UserForm1.Show vbModeless
Dim wv As stdWebView
Set wv = stdWebView.CreateFromFrame(UserForm1.FrameWeb)
wv.Navigate "https://example.com"
```

### Instance properties

#### `Html() As String` (Get)

Returns the document’s `document.documentElement.outerHTML`. Uses an internal synchronous script; requires `IsReady`.

#### `Html(ByVal rhs As String)` (Let)

Loads HTML into the view via WebView2’s string navigation (`NavigateToString`). Requires `IsReady`.

### Instance methods

#### `IsReady() As Boolean`

`True` when `CoreWebView2` is available. `Navigate`, `Html`, `JavaScriptRun`, and `JavaScriptRunSync` require a ready view.

#### `Quit()`

Tears down the controller reference and frees internal handler allocations. Safe to call when already shut down.

#### `Navigate(ByVal url As String)`

Navigates to a URL. Requires `IsReady`.

#### `Back()` / `Forward()`

History navigation implemented by running `history.back()` / `history.forward()` synchronously in the page.

#### `JavaScriptRunSync(ByVal script As String) As String`

Executes `script` in the page context and **blocks** until the result is delivered, pumping messages with `DoEvents`.

* Only **one** synchronous script may run at a time; a second call raises.
* The return value is the **JSON-encoded** result string from the WebView2 script API (e.g. quoted strings, `null`, numbers as JSON). Parse or unwrap as needed.

```vb
Debug.Print wv.JavaScriptRunSync("document.title")   ' e.g. returns JSON string including quotes
```

#### `JavaScriptRun(ByVal script As String, Optional ByVal callback As stdICallable = Nothing)`

Queues script execution without blocking. If `callback` is set, it is invoked with `RunEx(Array(errorCode, resultJson))` when execution completes.

```vb
wv.JavaScriptRun "console.log(1)", stdLambda.Create("Debug.Print $1, $2")  ' errorCode, resultJson
```

#### `AddHostObject(ByVal name As String, ByVal hostObject As Object)`

Injects a VBA COM object into JavaScript as `chrome.webview.hostObjects.<name>`.

* Requires `IsReady`.
* `name` must be non-empty.
* `hostObject` must be a non-`Nothing` object that is dispatchable (`IDispatch`).

```vb
' Class cBridge
' Public Function Echo(ByVal s As String) As String: Echo = "VBA:" & s: End Function

Dim bridge As New cBridge
wv.AddHostObject "bridge", bridge
```

JavaScript usage (async host object proxy):

```js
const value = await chrome.webview.hostObjects.bridge.Echo("hello");
console.log(value); // "VBA:hello"
```

#### `RemoveHostObject(ByVal name As String)`

Removes a previously injected object so `chrome.webview.hostObjects.<name>` is no longer available.

### Checklist (quick reference)

**Constructors**

* [X] `CreateFromHwnd(hwnd, OnReady?)`
* [X] `CreateFromFrame(frm, OnReady?)`

**Instance**

* [X] `IsReady()`
* [X] `Quit()`
* [X] `Navigate(url)`
* [X] `Back()` / `Forward()`
* [X] `Html` Get/Let
* [X] `JavaScriptRunSync(script)`
* [X] `JavaScriptRun(script, callback?)`
* [X] `AddHostObject(name, hostObject)`
* [X] `RemoveHostObject(name)`

### Protected / Friend API

`protCreate`, `protEnvCompleted`, `protCtrlCompleted`, and `protScriptCompleted` are **not** part of the public contract; they exist for the COM callback thunks. Do not call them from application code.

## stdVBA developer notes

* A per-instance user data folder is created under `%TEMP%` (`stdWebView_*`) for the WebView2 profile.
* If `CreateCoreWebView2EnvironmentWithOptions` fails, the error is raised with the HRESULT from the loader.
* `zzProtWebView_*` entry points must remain `Public` so thunk code can dispatch into the instance; they are not user APIs.
* Vtable offsets used in `stdWebView.cls` are verified against `ICoreWebView2Vtbl` in `WebView2.h` from the WebView2 SDK (`build/native/include/WebView2.h` in the NuGet package). Official header reference: [WebView2.h](https://github.com/MicrosoftEdge/WebView2Browser/blob/main/packages/Microsoft.Web.WebView2.1.0.2903.40/build/native/include/WebView2.h).
