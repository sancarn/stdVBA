# `stdHTTP`

## Index

* Introduction
* Spec
    * Constructors
        * `Create()`
        * `CreateOptions()`
        * `CreateHeaders()`
        * `CreateCookies()`
    * Instance Methods
        * `Await()`
        * `Get isFinished`
        * `Get ResponseStatus`
        * `Get ResponseStatusText`
        * `Get ResponseText`
        * `Get ResponseHeader(ByVal header as string)`
        * `Get ResponseHeaders`
    * Authorization
        * Basic
        * Windows


## Introduction

## Spec

### Constructors

#### `Create(ByVal sURL As String, Optional ByVal RequestMethod As String = "GET", Optional ByVal ThreadingStyle As EHTTPSynchronisity = HTTPSync, Optional ByVal options As Object = Nothing, Optional ByVal Authenticator As stdICallable = Nothing) as stdHTTP`

#### `CreateOptions(Optional Body As Variant = "", Optional Headers As Object<Dictionary> = Nothing, Optional Cookies As Object<Dictionary> = Nothing, Optional ByVal ContentTypeDefault As EHTTPContentType, Optional Insecure As Boolean = False, Optional EnableRedirects As Boolean = True, Optional ByVal TimeoutMS As Long = 5000, Optional ByVal AutoProxy As Boolean = True) as object<Dictionary>` 

Creates a HTTP options dictionary. Use this to specify the HTTP data-body, headers, cookies, timeout etc. All arguments are optional so it is advised to abuse VBA's named parameter access to bind variables as wanted. E.G.

```vb
Set o = stdHTTP.CreateOptions("hello body", Insecure:=True, EnableRedirects:=False)
```

> Remark: Options are loosely based on properties of [VBA-Web's Request](https://github.com/VBA-tools/VBA-Web/blob/master/src/WebRequest.cls) object and [JavaScript's `fetch` options](https://developer.mozilla.org/en-US/docs/Web/API/fetch).

A list of options and their uses can be found below

##### Body

Any body that you want to add to your request: this can be a `String`, `Long`, `Double`, `Array` or any variable that can be passed to `VARIANT`.

> Note: Data is passed as-is without any form of serialisation.

> Note: A request using the `GET` or `HEAD` methods cannot have a body.

##### Headers

Any headers you want to add to your request, contained within a `Dictionary<string,string>` object. Can also use `stdHTTP.CreateHeaders()`

<!-- TODO: > Note that [some names are forbidden](...). -->

##### Cookies

Any cookies you want to add to your request for authentication purposes or otherwise. to be a `Dictionary<string,string>` object. Can also use `stdHTTP.CreateCookies()`.

##### ContentTypeDefault

Any combination of the following values:

```vb
ContentType_HTML = 1
ContentType_Plain = 2
ContentType_JSON = 4
ContentType_XML = 8
ContentType_CSV = 16
ContentType_ZIP = 32
ContentType_Binary = 64
```

This is a non-exhaustive list and is totally overwritten by the `Content-Type` header if provided. However this contains some types so you don't have to remember the specific content type string.

##### Insecure

If set to `true`:

* SSL Errors will be ignored
* Certificates will not be required
* HTTPS -> HTTP redirects will be allowed

The opposite if set to `false`.

##### EnableRedirects

If set to `true` HTTP will follow any redirects, else it will error on redirect.

##### TimeoutMS

The timeout for requests measured in milliseconds.

<!-- TODO: sometimes doesn't seem to work properly? -->

##### AutoProxy

Whether to automatically search for proxies for the current user. Useful especially in businesses.

#### CreateHeaders(ParamArray v())

Create a dictionary of header keys and values.


