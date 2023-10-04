Attribute VB_Name = "stdHTTPTests"
'@lang VBA

Public Sub TestMain()
  Test.Topic "stdHTTP"
  
  Dim r As stdHTTP
  
  'Test GET
  Set r = stdHTTP.Create("https://postman-echo.com/get")
  Test.Assert "GET Test 1) Status Complete", r.ResponseStatus = 200
  Test.Assert "GET Test 1) URL as expected", r.ResponseText Like "{*""url"":""https://postman-echo.com/get""*}"
  Test.Assert "GET Test 1) no args as expected", r.ResponseText Like "{*""args"":{}*}"
  
  'Test GET Async
  Set r = stdHTTP.Create("https://postman-echo.com/get", ThreadingStyle:=HTTPAsync)
  Test.Assert "GET Test 2) Async Immediate Status == 0", r.ResponseStatus = 0
  Test.Assert "GET Test 2) Async Immediate StatusText == Awaiting Response", r.ResponseStatusText = "Awaiting Response"
  With r.Await()
    Test.Assert "GET Test 2) Async Final Status == 200", r.ResponseStatus = 200
    Test.Assert "GET Test 2) Async Final StatusText == OK", r.ResponseStatusText = "OK"
  End With
  
  'Test GET With headers
  Set r = stdHTTP.Create("https://postman-echo.com/get", options:=stdHTTP.CreateOptions(Headers:=stdHTTP.CreateHeaders("XXX-CustomHeader", 1, "user-agent", "stdHTTP::Test")))
  Test.Assert "GET Test 3) Headers - Custom Header added", r.ResponseText Like "{*""xxx-customheader"":""1""*}"
  Test.Assert "GET Test 3) Headers - Custom Header added", r.ResponseText Like "{*""user-agent"":""stdHTTP::Test""*}"
  
  'Test POST
  Set r = stdHTTP.Create("https://postman-echo.com/post", "POST", options:=stdHTTP.CreateOptions("hello body"))
  Test.Assert "POST Test 1) Status Complete", r.ResponseStatus = 200
  Test.Assert "POST Test 1) URL as expected", r.ResponseText Like "{*""url"":""https://postman-echo.com/post""*}"
  Test.Assert "POST Test 1) Data/Body as provided", r.ResponseText Like "{*""data"":""hello body""*}"
  
  
  'Test basic authenticator:
  Set r = stdHTTP.Create("https://postman-echo.com/basic-auth")
  Test.Assert "AUTH BASIC 1) No user/pass == rejection", r.ResponseStatus >= 400
  Set r = stdHTTP.Create("https://postman-echo.com/basic-auth", Authenticator:=stdCallback.CreateFromModule("stdHTTPAuthenticators", "HttpBasicAuthenticator").Bind("postman", "password"))
  Test.Assert "AUTH BASIC 2) Correct user/pass == OK", r.ResponseStatus = 200
  
  'Test cookies
  Set r = stdHTTP.Create("https://postman-echo.com/cookies", options:=stdHTTP.CreateOptions(Cookies:=stdHTTP.CreateHeaders("cookieKey", "cookieVal", "cookie2", "cookieVal2")))
  Test.Assert "COOKIE 1) Has first cookie", r.ResponseText Like "{*""cookieKey"":""cookieVal""*}"
  Test.Assert "COOKIE 2) Has both cookies", r.ResponseText Like "{*""cookie2"":""cookieVal2""*}"
  
  'Test response headers
  Set r = stdHTTP.Create("https://postman-echo.com/response-headers?foo=bar")
  Test.Assert "GET 4) Response header get", r.ResponseHeader("foo") = "bar"
  
  'TBC: Find a good test for Insecure
  'Set r = stdHTTP.Create("https://postman-echo.com/response-headers", options:=stdHTTP.CreateOptions(Insecure:=True))
  
  'TBC: Can we work with Streamed responses?
  'set r = stdHTTP.Create("https://postman-echo.com/stream/5")
  
  'TBC: Force timeout
  'Set r = stdHTTP.Create("https://postman-echo.com/delay/2", options:=stdHTTP.CreateOptions(TimeoutMS:=1000))
  
  'TBC: Gzip compressed response https://postman-echo.com/gzip
  
  'TBC: Deflate compressed response https://postman-echo.com/deflate
  Debug.Assert False
End Sub

