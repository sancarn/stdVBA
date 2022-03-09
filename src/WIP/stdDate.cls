VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Class that wraps standard VBA DateTime functionality:
'* Create from double
'* Create from Year, Month, Day, Hour, Min, Second
'* Create from POSIX datetime
'* [DEFAULT PROP] Value; get and let as Double.
'* [PROP] Seconds; get
'* [PROP] Minutes; get
'* [PROP] Hours; get
'* [PROP] Days; get
'* [PROP] Months; get
'* [PROP] Years; get
'* [PROP] Weekday; get
'* [PROP] WeekdayName; get
'* [PROP] MonthName; get
'* [PROP] Country; get and let   - Timezone - Still not sure whether this will work well enough.
'* [METHOD] Parse(sVal, sFormat) - For user defined parsing
'* [METHOD] Format(sFormat)      - For user defined parsing
'* [METHOD] ToString()           - Do yyyymmddThhMMss format
'* [METHOD] isLeap()             - Is this date a leap year?
'
'Usage examples:
'  Format cell A1 in dd-mm-yyyy format.
'    Debug.Print stdDate.Create(Range("A1")).Format("dd-mm-yyyy")
'
'  Format yyyy-mm-dd hh-mm-ss as dd/mm/yy hh:mm:ss 
'    Debug.Print stdDate.CreateFromParse("2012-02-01 10-11-35","yyyy-mm-dd hh-mm-ss").Format("dd/mm/yy hh:mm:ss")
'

Private pInitialised As Boolean
Private pValue as Date
Private pCountry as String
Private pFirstDayOfWeek as VbDayOfWeek

'TODO: Determine whether these are actually useful or not
Public daysPerMonth As Object
Public TimeZones As Object
Public TimeZoneDescs As Object
Public MonthsDict As Object
Public MonthsShortDict As Object

Friend Sub Init(val as Date)
  if not pInitialised then
    Country = getGeoCode()
    pValue = val
    pInitialised = true
  else
    'Call STD.Errors.Raise("STD.Class",1)
  end if
End Sub

Public Function Create(Optional val as Date) as stdDate
  if not pInitialised then
    set Create = new stdDate
    Call Create.Init(val)
  else
    'Call STD.Errors.Raise("STD.Class",2)
  end if
End Function

'
Public Function CreateFromUnits(Optional ByVal year As Integer = 0, Optional ByVal month As Integer = 0, Optional ByVal day As Integer = 0, Optional ByVal Hour As Integer = 0, Optional ByVal Minute As Integer = 0, Optional ByVal Second As Integer = 0) As stdDate
  if not pInitialised then
    set CreateFromUnits = Create(DateSerial(Year,Month,Day) + TimeSerial(Hour,Minute,Second))
  else
    'Call STD.Errors.Raise("STD.Class",2)
  end if
End Function

'Tries to emulate: https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html
'EEE  = Mon|Tue|Wd|Thu|Fri
'E    = Monday|Tuesday|Wednesday|Thursday|Friday
'For hh:mm:ss dd/MM/yyyy = 21:22:23 01/02/2019
'yyyy = 2019     (?<Y4>\d{2}\d{2}?)
'yy   = 19       (?<Y2>\d{1,2})
'MMMM = February (?<M4>January|February|March|...)
'MMM  = Feb      (?<M3>Jan|Feb|Mar|...)
'MM   = 02       (?<M2>\d{1,2})
'dd   = 01       (?<D2>\d{1,2})
'hh   = 21       (?<H2>\d{1,2})
'kk   = 09       (?<K2>\d{1,2})
'a    = pm       (?<A1>am|pm)
'mm   = 22       (?<MM>\d{1,2})
'ss   = 23       (?<S2>\d{1,2})
'zzzz = Grenich Mean Time (?<Z4>Grenich Mean Time|Pacific Standard Time|...)
'zzz  = GMT      (?<Z3>GMT|PST|...)
'All other values in java are unsupported
'@param 
'@returns {stdDate} Parsed date object
Public Function CreateFromParse(ByVal Value As String, Optional strFormat As String = "") As stdDate
  Dim d as long,m as long,y as long,h as long,mn as long,s as long

  'Common formats:
  If strFormat = "" Or strFormat like "dd[./-_]mm[./-_]yyyy" Then
    'Parse assuming dd/mm/yyyy
    d = CInt(Mid(Value, 1, 2))
    m = CInt(Mid(Value, 4, 2))
    y = CInt(Mid(Value, 7, 4))
  ElseIf strFormat like "mm[./-_]dd[./-_]yyyy" Then
    d = CInt(Mid(Value, 4, 2))
    m = CInt(Mid(Value, 1, 2))
    y = CInt(Mid(Value, 7, 4))
    
  ElseIf strFormat = "UTC" Then
    'yyyy-MM-ddThh:mm:ss.llll  (llll = milliseconds, T = "T")
    ms = CInt(Mid(Value, InStr(1, Value, ".")))
    s  = CInt(Mid(Value, 18, 2))
    mn = CInt(Mid(Value, 15, 2))
    h  = CInt(Mid(Value, 12, 2))
    d  = CInt(Mid(Value, 9, 2))
    m  = CInt(Mid(Value, 6, 2))
    y  = CInt(Mid(Value, 1, 4))
    
  ElseIf strFormat = "dd/mm/yyyy hh:mm:ss" Then
    d = CInt(Mid(Value, 1, 2))
    m = CInt(Mid(Value, 4, 2))
    y = CInt(Mid(Value, 7, 4))
    h   = CInt(Mid(Value, 12, 2))
    mn = CInt(Mid(Value, 15, 2))
    s = CInt(Mid(Value, 18, 2))
  Else
    
    'USE REGEX:
    'strFormat = Replace(strFormat, "EEE" , "(?<E3>Mon|Tue|Wd|Thu|Fri)")
    'strFormat = Replace(strFormat, "EE"  , "(?<E2>Monday|Tuesday|Wednesday|Thursday|Friday)")
    'strFormat = Replace(strFormat, "yyyy", "(?<Y4>\d{4})")
    'strFormat = Replace(strFormat, "yy"  , "(?<Y2>\d{2})")
    'strFormat = Replace(strFormat, "mmmm", "(?<M4>January|February|March|April|May|June|July|August|September|October|November|December)")
    'strFormat = Replace(strFormat, "mmm" , "(?<M3>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)")
    'strFormat = Replace(strFormat, "mm"  , "(?<M2>\d{1,2})")
    'strFormat = Replace(strFormat, "dd"  , "(?<D2>\d{1,2})")
    'strFormat = Replace(strFormat, "hh"  , "(?<hh>\d{1,2})")
    'strFormat = Replace(strFormat, "kk"  , "(?<kk>\d{1,2})")
    'strFormat = Replace(strFormat, "aa"  , "(?<aa>am|pm)")
    'strFormat = Replace(strFormat, "MM"  , "(?<MM>\d{1,2})")
    'strFormat = Replace(strFormat, "ss"  , "(?<ss>\d{1,2})")
    'strFormat = Replace(strFormat, "zzzz", "(?<Z4>" & TimeZoneDescs.keys().Join("|") & ")")
    'strFormat = Replace(strFormat, "zzz" , "(?<Z3>" & TimeZones.keys().Join("|") & ")")
    
    strFormat = Replace(strFormat, "EEE", "(?<E3>Mon|Tue|Wd|Thu|Fri)")
    strFormat = Replace(strFormat, "EE", "(?<E2>Monday|Tuesday|Wednesday|Thursday|Friday)")
    strFormat = Replace(strFormat, "yyyy", "(?<Y4>\d{4})")
    strFormat = Replace(strFormat, "yy", "(?<Y2>\d{2})")
    strFormat = Replace(strFormat, "mmmm", "(?<M4>January|February|March|April|May|June|July|August|September|October|November|December)")
    strFormat = Replace(strFormat, "mmm", "(?<M3>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)")
    strFormat = Replace(strFormat, "mm", "(?<M2>\d{1,2})")
    strFormat = Replace(strFormat, "dd", "(?<D2>\d{1,2})")
    strFormat = Replace(strFormat, "hh", "(?<hh>\d{1,2})")
    strFormat = Replace(strFormat, "kk", "(?<kk>\d{1,2})")
    strFormat = Replace(strFormat, "aa", "(?<aa>am|pm)")
    strFormat = Replace(strFormat, "MM", "(?<mm>\d{1,2})")
    strFormat = Replace(strFormat, "ss", "(?<ss>\d{1,2})")
    
    'If it has timezones
    Dim hasTZs As Boolean
    hasTZs = InStr(1, strFormat, "zzz")
    If hasTZs Then
      strFormat = Replace(strFormat, "zzzz", "(?<Z4>" & TimeZoneDescs.keys().Join("|") & ")")
      strFormat = Replace(strFormat, "zzz", "(?<Z3>" & TimeZones.keys().Join("|") & ")")
    End If
    
    'Parse using regex:
    Dim match As Object
    Set match = Me.RegexMatch(Value, strFormat, "i")
    
    Dim century As Long: century = CInt(year(Now()) / 100) * 100
    
    'Y2 is an odd format, e.g. If the current year is 2020 and you see 01-01-21 written down,
    'you will likely thing this is the 1st Jan 2021. But if you see 01-01-91, you are more likely
    'to see this as 1991 than 01-01-2091.
    'The y2Error allows us to detect all dates whos year is less than the current decade + y2Error
    'as belonging to that century.
    Dim y2Error As Long: y2Error = 5
    
    'Parse datetime
    If match("Y4") Then y = CInt(match("Y4"))
    If match("Y2") Then y = IIf(CInt(match("Y2")) < (year(Now()) - century + y2Error), century + CInt(match("Y2")), century - 100 + CInt(match("Y2")))
    If match("M4") Then m = stdDate.MonthsDict(match("M4"))
    If match("M3") Then m = stdDate.MonthsShortDict(match("M3"))
    If match("M2") Then m = CInt(match("M2"))
    If match("D2") Then d = CInt(match("D2"))
    If match("hh") Then h = CInt(match("hh"))
    If match("MM") Then mn = CInt(match("MM"))
    If match("ss") Then s = CInt(match("ss"))
    
    'Special:
    If match("kk") And match("aa") Then h = IIf(match("aa") = "am", 0, 12) + CInt(match("kk"))
    
    If hasTZs Then
      Err.Raise 0, "stdDate::Parse", "Error: Timezones Unimplemented. Waiting for swith to CoreTime"
      'Call iniTimezones
      'Dim offsetCoreTime As Double
      'If match("zzzz") Then offsetCoreTime = match("zzzz") * foo
      'If match("zzz") Then offsetCoreTime = match("zzz") * foo
      'pCoreTime = pCoreTime - offsetCoreTime
    End If
  End If
  
  'Create return value from CreateFromUnits
  set CreateFromParse = stdDate.CreateFromUnits(y, m, d, h, mn, s)
End Function

'TODO:
'Create date from POSIX date
'@param {Double} Date to convert
'@returns {stdDate} Date object
'@remark EPOCH == 00:00:00 01/01/1970
'Public Function CreateFromPOSIX(ByVal date as double) as stdDate
'  Err.Raise 0, "stdDate::Parse", "Currently WIP not yet functional"
'  
'  'Get value as double
'  Dim dbl As Double
'  dbl = CDbl(Value)
'  
'  
'  ms = Int(dmod(dbl * 1000, 1000))
'  s = Int(dmod(dbl, 60))
'  mn = Int(dmod(dbl / 60, 60))
'  h = Int(dmod(dbl / 60 / 60, 24))
'  d = Int(dmod(dbl / 60 / 60 / 24, 365))
'  'm '=...?
'  'y '=...
'End Function


Public Property Get Value() as Double
Attribute Value.Vb_UserMemId = 0
  if pInitialised then
    Value = pValue
  else
    'Call STD.Errors.Raise("STD.Class",4)
  end if
End Property
Public Property Let Value(v as Double)
  if pInitialised then
    pValue = v
  else
    'Call STD.Errors.Raise("STD.Class",5)
  end if
End Property

'Common properties getters - Not sure if these need setters yet...:
Public Property Get Seconds() as Long
  If pInitialised then
    Seconds = VBA.Second(pValue)
  Else
    'Call STD.Errors.Raise("STD.Class",4)
  End if
End Property
Public Property Get Minutes() as Long
  If pInitialised then
    Minutes = VBA.Minute(pValue)
  Else
    'Call STD.Errors.Raise("STD.Class",4)
  End if
End Property
Public Property Get Hours() as Long
  If pInitialised then
    Hours = VBA.Hour(pValue)
  Else
    'Call STD.Errors.Raise("STD.Class",4)
  End if
End Property
Public Property Get Days() as Long
  If pInitialised then
    Days = VBA.Day(pValue)
  Else
    'Call STD.Errors.Raise("STD.Class",4)
  End if
End Property
Public Property Get Months() as Long
  If pInitialised then
    Months = VBA.Month(pValue)
  Else
    'Call STD.Errors.Raise("STD.Class",4)
  End if
End Property
Public Property Get Years() as Long
  If pInitialised then
    Years = VBA.Year(pValue)
  Else
    'Call STD.Errors.Raise("STD.Class",4)
  End if
End Property

'Get integer position of day within week. E.G. 1-Monday, 2-Tuesday, 3-Wednesday, ...
Public Property Get Weekday(Optional ByVal FirstDayOfWeek as VbDayOfWeek) as Long
  If pInitialised then
    If Not FirstDayOfWeek then FirstDayOfWeek = pFirstDayOfWeek
    Weekday = VBA.Weekday(pValue,FirstDayOfWeek)
  Else
    'Call STD.Errors.Raise("STD.Class",4)
  End If
End Property
Public Property Get WeekdayName(Optional ByVal Abbreviate as Boolean = false,Optional ByVal FirstDayOfWeek as VbDayOfWeek) as Long
  If pInitialised then
    If Not FirstDayOfWeek then FirstDayOfWeek = pFirstDayOfWeek
    WeekdayName = VBA.WeekdayName(Weekday,Abbreviate,FirstDayOfWeek)
  Else
    'Call STD.Errors.Raise("STD.Class",4)
  End If
End Property
Public Property Get MonthName(Optional Abbreviate as boolean = false)
  if pInitialised then
    MonthName = VBA.MonthName(VBA.Month(pValue),Abbreviate)
  Else
    'Call STD.Errors.Raise("STD.Class",4)
  End If
End Property

'Set and Get country used in Getters for Weekday and Weekname
Public Property Get Country() as string
  Country = pCountry
End Property
Public Property Let Country(s as string)
  pCountry = s
  select case s
    case "US":
      pFirstDayOfWeek = vbSunday
    case "UK":
      pFirstDayOfWeek = vbMonday
  end select
End Property

Public Function ToString()
  If pInitialised then
    ToString = VBA.Format(pValue,"yyyymmddhhmmss")
  Else
    'Call STD.Errors.Raise("STD.Class",3)
  End If
End Function

Public Property Get isLeap() As Boolean
  Dim yr as long
  yr = Years

  'Excel incorrectly treats 1900 as a leap year. See https://en.wikipedia.org/wiki/Leap_year_bug
  If yr = 1900 Then isLeap = True: Exit Function
  
  'Regular leap years:
  isLeap = ((yr Mod 4 = 0 And yr Mod 100 <> 0) Or yr Mod 400 = 0)
End Property

'Example format strings:
'd/m/y               --> 1/2/9
'dd/mm/yy            --> 01/02/19
'dd/mm/yyyy          --> 01/02/2019
'mm/dd/yyyy          --> 02/01/2019
'DD D MMMM yyyy      --> Monday 1st February 2019
'dd/mm/yyyy hh:MM:ss --> 01/02/2019 12:01:00
Public Function Format(strFormat As String) As String
  If initialised Then
    Format = "...todo..."
  Else
    Err.Raise 0, "Cannot format uninitialised date"
  End If
End Function




'Private date funcs
Private Function p_isLeap(ByVal yr As Integer) As Boolean
  'Excel incorrectly treats 1900 as a leap year. See https://en.wikipedia.org/wiki/Leap_year_bug
  If yr = 1900 Then p_isLeap = True: Exit Function
  
  p_isLeap = ((yr Mod 4 = 0 And yr Mod 100 <> 0) Or yr Mod 400 = 0)
End Function

Private Function DaysBetween(d1 As stdDate, d2 As stdDate) As Long
  'Subtract d2 from d1. I.E. D1-D2
  Dim Days1, Days2 As Double
  Dim Factor As Double
  Factor = 365.25 '365.2425
  
  'Calculate total days in d1 and d2.
  'Note: In February, 31 days (January have gone by) Thus daysPerMonth("Cumulative") has 2 zeros at the start of the array
  Days1 = d1.days + stdDate.daysPerMonth("Cumulative" & IIf(d1.isLeap, "Leap", ""))(d1.months) + d1.years * Factor
  Days2 = d2.days + stdDate.daysPerMonth("Cumulative" & IIf(d2.isLeap, "Leap", ""))(d2.months) + d2.years * Factor
  
  'Subtract the 2 day counts. Also include the fake leap year in 1900 created by Excel
  DaysBetween = Days1 - Days2 + IIf(d2.years = 1900 And d1.years <> 1900, 1, 0)
End Function

Function RoundUp(ByVal Value As Double)
    If Int(Value) = Value Then
        RoundUp = Value
    Else
        RoundUp = Int(Value) + 1
    End If
End Function

'INITIALISATION:
Private Sub Class_Initialize()
  Call iniDictionary
End Sub

Private Sub iniDictionary()
  if me is stdDate then
    Set daysPerMonth = CreateObject("Scripting.Dictionary")
    daysPerMonth(False) = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    daysPerMonth(True) = Array(0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    daysPerMonth("Cumulative") = Array(0, 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365)
    daysPerMonth("CumulativeLeap") = Array(0, 0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366)
    
    'Months hash table
    Set MonthsDict = CreateObject("Scripting.Dictionary")
    MonthsDict("January") = 1
    MonthsDict("February") = 2
    MonthsDict("March") = 3
    MonthsDict("April") = 4
    MonthsDict("May") = 5
    MonthsDict("June") = 6
    MonthsDict("July") = 7
    MonthsDict("August") = 8
    MonthsDict("September") = 9
    MonthsDict("October") = 10
    MonthsDict("November") = 11
    MonthsDict("December") = 12
    
    'Short months hash table
    Set MonthsShortDict = CreateObject("Scripting.Dictionary")
    MonthsShortDict("Jan") = 1
    MonthsShortDict("Feb") = 2
    MonthsShortDict("Mar") = 3
    MonthsShortDict("Apr") = 4
    MonthsShortDict("May") = 5
    MonthsShortDict("Jun") = 6
    MonthsShortDict("Jul") = 7
    MonthsShortDict("Aug") = 8
    MonthsShortDict("Sep") = 9
    MonthsShortDict("Oct") = 10
    MonthsShortDict("Nov") = 11
    MonthsShortDict("Dec") = 12
  end if
End Sub


Private Sub iniTimezones()
  'This slows down a lot of stuff so make sure it only runs on the static class
  if Me is stdDate then
    'Timezones:
    Set TimeZones = CreateObject("Scripting.Dictionary")
    Set TimeZoneDescs = CreateObject("Scripting.Dictionary")
    
    'Time zone codes
    TimeZones("ACDT") = 10.5
    TimeZones("ACST") = 9.5
    TimeZones("ACT") = -5
    TimeZones("ACT") = 6.5
    TimeZones("ACWST") = 8.75
    TimeZones("ADT") = -3
    TimeZones("AEDT") = 11
    TimeZones("AEST") = 10
    TimeZones("AFT") = 4.5
    TimeZones("AKDT") = -8
    TimeZones("AKST") = -9
    TimeZones("AMST") = -3
    TimeZones("AMT") = -4
    TimeZones("AMT") = 4
    TimeZones("ART") = -3
    TimeZones("AST") = 3
    TimeZones("AST") = -4
    TimeZones("AWST") = 8
    TimeZones("AZOST") = 0
    TimeZones("AZOT") = -1
    TimeZones("AZT") = 4
    TimeZones("BDT") = 8
    TimeZones("BIOT") = 6
    TimeZones("BIT") = -12
    TimeZones("BOT") = -4
    TimeZones("BRST") = -2
    TimeZones("BRT") = -3
    TimeZones("BST") = 6
    TimeZones("BST") = 11
    TimeZones("BST") = 1
    TimeZones("BTT") = 6
    TimeZones("CAT") = 2
    TimeZones("CCT") = 6.5
    TimeZones("CDT") = -5
    TimeZones("CDT") = -4
    TimeZones("CEST") = 2
    TimeZones("CET") = 1
    TimeZones("CHADT") = 13.75
    TimeZones("CHAST") = 12.75
    TimeZones("CHOT") = 8
    TimeZones("CHOST") = 9
    TimeZones("CHST") = 10
    TimeZones("CHUT") = 10
    TimeZones("CIST") = -8
    TimeZones("CIT") = 8
    TimeZones("CKT") = -10
    TimeZones("CLST") = -3
    TimeZones("CLT") = -4
    TimeZones("COST") = -4
    TimeZones("COT") = -5
    TimeZones("CST") = -6
    TimeZones("CST") = 8
    TimeZones("CST") = -5
    TimeZones("CT") = 8
    TimeZones("CVT") = -1
    TimeZones("CWST") = 8.75
    TimeZones("CXT") = 7
    TimeZones("DAVT") = 7
    TimeZones("DDUT") = 10
    TimeZones("DFT") = 1
    TimeZones("EASST") = -5
    TimeZones("EAST") = -6
    TimeZones("EAT") = 3
    TimeZones("ECT") = -4
    TimeZones("ECT") = -5
    TimeZones("EDT") = -4
    TimeZones("EEST") = 3
    TimeZones("EET") = 2
    TimeZones("EGST") = 0
    TimeZones("EGT") = -1
    TimeZones("EIT") = 9
    TimeZones("EST") = -5
    TimeZones("FET") = 3
    TimeZones("FJT") = 12
    TimeZones("FKST") = -3
    TimeZones("FKT") = -4
    TimeZones("FNT") = -2
    TimeZones("GALT") = -6
    TimeZones("GAMT") = -9
    TimeZones("GET") = 4
    TimeZones("GFT") = -3
    TimeZones("GILT") = 12
    TimeZones("GIT") = -9
    TimeZones("GMT") = 0
    TimeZones("GST") = -2
    TimeZones("GST") = 4
    TimeZones("GYT") = -4
    TimeZones("HDT") = -9
    TimeZones("HAEC") = 2
    TimeZones("HST") = -10
    TimeZones("HKT") = 8
    TimeZones("HMT") = 5
    TimeZones("HOVST") = 8
    TimeZones("HOVT") = 7
    TimeZones("ICT") = 7
    TimeZones("IDLW") = -12
    TimeZones("IDT") = 3
    TimeZones("IOT") = 3
    TimeZones("IRDT") = 4.5
    TimeZones("IRKT") = 8
    TimeZones("IRST") = 3.5
    TimeZones("IST") = 5.5
    TimeZones("IST") = 1
    TimeZones("IST") = 2
    TimeZones("JST") = 9
    TimeZones("KALT") = 2
    TimeZones("KGT") = 6
    TimeZones("KOST") = 11
    TimeZones("KRAT") = 7
    TimeZones("KST") = 9
    TimeZones("LHST") = 10.5
    TimeZones("LHST") = 11
    TimeZones("LINT") = 14
    TimeZones("MAGT") = 12
    TimeZones("MART") = -8.5
    TimeZones("MAWT") = 5
    TimeZones("MDT") = -6
    TimeZones("MET") = 1
    TimeZones("MEST") = 2
    TimeZones("MHT") = 12
    TimeZones("MIST") = 11
    TimeZones("MIT") = -8.5
    TimeZones("MMT") = 6.5
    TimeZones("MSK") = 3
    TimeZones("MST") = 8
    TimeZones("MST") = -7
    TimeZones("MUT") = 4
    TimeZones("MVT") = 5
    TimeZones("MYT") = 8
    TimeZones("NCT") = 11
    TimeZones("NDT") = -1.5
    TimeZones("NFT") = 11
    TimeZones("NPT") = 5.75
    TimeZones("NST") = -2.5
    TimeZones("NT") = -2.5
    TimeZones("NUT") = -11
    TimeZones("NZDT") = 13
    TimeZones("NZST") = 12
    TimeZones("OMST") = 6
    TimeZones("ORAT") = 5
    TimeZones("PDT") = -7
    TimeZones("PET") = -5
    TimeZones("PETT") = 12
    TimeZones("PGT") = 10
    TimeZones("PHOT") = 13
    TimeZones("PHT") = 8
    TimeZones("PKT") = 5
    TimeZones("PMDT") = -2
    TimeZones("PMST") = -3
    TimeZones("PONT") = 11
    TimeZones("PST") = -8
    TimeZones("PST") = 8
    TimeZones("PYST") = -3
    TimeZones("PYT") = -4
    TimeZones("RET") = 4
    TimeZones("ROTT") = -3
    TimeZones("SAKT") = 11
    TimeZones("SAMT") = 4
    TimeZones("SAST") = 2
    TimeZones("SBT") = 11
    TimeZones("SCT") = 4
    TimeZones("SDT") = -10
    TimeZones("SGT") = 8
    TimeZones("SLST") = 5.5
    TimeZones("SRET") = 11
    TimeZones("SRT") = -3
    TimeZones("SST") = -11
    TimeZones("SST") = 8
    TimeZones("SYOT") = 3
    TimeZones("TAHT") = -10
    TimeZones("THA") = 7
    TimeZones("TFT") = 5
    TimeZones("TJT") = 5
    TimeZones("TKT") = 13
    TimeZones("TLT") = 9
    TimeZones("TMT") = 5
    TimeZones("TRT") = 3
    TimeZones("TOT") = 13
    TimeZones("TVT") = 12
    TimeZones("ULAST") = 9
    TimeZones("ULAT") = 8
    TimeZones("UTC") = 0
    TimeZones("UYST") = -2
    TimeZones("UYT") = -3
    TimeZones("UZT") = 5
    TimeZones("VET") = -4
    TimeZones("VLAT") = 10
    TimeZones("VOLT") = 4
    TimeZones("VOST") = 6
    TimeZones("VUT") = 11
    TimeZones("WAKT") = 12
    TimeZones("WAST") = 2
    TimeZones("WAT") = 1
    TimeZones("WEST") = 1
    TimeZones("WET") = 0
    TimeZones("WIT") = 7
    TimeZones("WST") = 8
    TimeZones("YAKT") = 9
    TimeZones("YEKT") = 5
    
    'Time zone descriptions
    TimeZoneDescs("Australian Central Daylight Savings Time") = 10.5
    TimeZoneDescs("Australian Central Standard Time") = 9.5
    TimeZoneDescs("Acre Time") = -5
    TimeZoneDescs("ASEAN Common Time") = 6.5
    TimeZoneDescs("Australian Central Western Standard Time�(unofficial)") = 8.75
    TimeZoneDescs("Atlantic Daylight Time") = -3
    TimeZoneDescs("Australian Eastern Daylight Savings Time") = 11
    TimeZoneDescs("Australian Eastern Standard Time") = 10
    TimeZoneDescs("Afghanistan Time") = 4.5
    TimeZoneDescs("Alaska Daylight Time") = -8
    TimeZoneDescs("Alaska Standard Time") = -9
    TimeZoneDescs("Amazon Summer Time�(Brazil)[1]") = -3
    TimeZoneDescs("Amazon Time�(Brazil)[2]") = -4
    TimeZoneDescs("Armenia Time") = 4
    TimeZoneDescs("Argentina Time") = -3
    TimeZoneDescs("Arabia Standard Time") = 3
    TimeZoneDescs("Atlantic Standard Time") = -4
    TimeZoneDescs("Australian Western Standard Time") = 8
    TimeZoneDescs("Azores Summer Time") = 0
    TimeZoneDescs("Azores Standard Time") = -1
    TimeZoneDescs("Azerbaijan Time") = 4
    TimeZoneDescs("Brunei Time") = 8
    TimeZoneDescs("British Indian Ocean Time") = 6
    TimeZoneDescs("Baker Island Time") = -12
    TimeZoneDescs("Bolivia Time") = -4
    TimeZoneDescs("Bras�lia Summer Time") = -2
    TimeZoneDescs("Brasilia Time") = -3
    TimeZoneDescs("Bangladesh Standard Time") = 6
    TimeZoneDescs("Bougainville Standard Time[3]") = 11
    TimeZoneDescs("British Summer Time�(British Standard Time�from Feb 1968 to Oct 1971)") = 1
    TimeZoneDescs("Bhutan Time") = 6
    TimeZoneDescs("Central Africa Time") = 2
    TimeZoneDescs("Cocos Islands Time") = 6.5
    TimeZoneDescs("Central Daylight Time�(North America)") = -5
    TimeZoneDescs("Cuba Daylight Time[4]") = -4
    TimeZoneDescs("Central European Summer Time�(Cf. HAEC)") = 2
    TimeZoneDescs("Central European Time") = 1
    TimeZoneDescs("Chatham Daylight Time") = 13.75
    TimeZoneDescs("Chatham Standard Time") = 12.75
    TimeZoneDescs("Choibalsan Standard Time") = 8
    TimeZoneDescs("Choibalsan Summer Time") = 9
    TimeZoneDescs("Chamorro Standard Time") = 10
    TimeZoneDescs("Chuuk Time") = 10
    TimeZoneDescs("Clipperton Island Standard Time") = -8
    TimeZoneDescs("Central Indonesia Time") = 8
    TimeZoneDescs("Cook Island Time") = -10
    TimeZoneDescs("Chile Summer Time") = -3
    TimeZoneDescs("Chile Standard Time") = -4
    TimeZoneDescs("Colombia Summer Time") = -4
    TimeZoneDescs("Colombia Time") = -5
    TimeZoneDescs("Central Standard Time�(North America)") = -6
    TimeZoneDescs("China Standard Time") = 8
    TimeZoneDescs("Cuba Standard Time") = -5
    TimeZoneDescs("China Time") = 8
    TimeZoneDescs("Cape Verde Time") = -1
    TimeZoneDescs("Central Western Standard Time�(Australia) unofficial") = 8.75
    TimeZoneDescs("Christmas Island Time") = 7
    TimeZoneDescs("Davis Time") = 7
    TimeZoneDescs("Dumont d'Urville Time") = 10
    TimeZoneDescs("AIX-specific equivalent of�Central European Time[NB 1]") = 1
    TimeZoneDescs("Easter Island Summer Time") = -5
    TimeZoneDescs("Easter Island Standard Time") = -6
    TimeZoneDescs("East Africa Time") = 3
    TimeZoneDescs("Eastern Caribbean Time�(does not recognise DST)") = -4
    TimeZoneDescs("Ecuador Time") = -5
    TimeZoneDescs("Eastern Daylight Time�(North America)") = -4
    TimeZoneDescs("Eastern European Summer Time") = 3
    TimeZoneDescs("Eastern European Time") = 2
    TimeZoneDescs("Eastern Greenland Summer Time") = 0
    TimeZoneDescs("Eastern Greenland Time") = -1
    TimeZoneDescs("Eastern Indonesian Time") = 9
    TimeZoneDescs("Eastern Standard Time (North America)") = -5
    TimeZoneDescs("Further-eastern European Time") = 3
    TimeZoneDescs("Fiji Time") = 12
    TimeZoneDescs("Falkland Islands Summer Time") = -3
    TimeZoneDescs("Falkland Islands Time") = -4
    TimeZoneDescs("Fernando de Noronha Time") = -2
    TimeZoneDescs("Gal�pagos Time") = -6
    TimeZoneDescs("Gambier Islands Time") = -9
    TimeZoneDescs("Georgia Standard Time") = 4
    TimeZoneDescs("French Guiana Time") = -3
    TimeZoneDescs("Gilbert Island Time") = 12
    TimeZoneDescs("Gambier Island Time") = -9
    TimeZoneDescs("Greenwich Mean Time") = 0
    TimeZoneDescs("South Georgia and the South Sandwich Islands Time") = -2
    TimeZoneDescs("Gulf Standard Time") = 4
    TimeZoneDescs("Guyana Time") = -4
    TimeZoneDescs("Hawaii�Aleutian Daylight Time") = -9
    TimeZoneDescs("Heure Avanc�e d'Europe Centrale�French-language name for CEST") = 2
    TimeZoneDescs("Hawaii�Aleutian Standard Time") = -10
    TimeZoneDescs("Hong Kong Time") = 8
    TimeZoneDescs("Heard and McDonald Islands�Time") = 5
    TimeZoneDescs("Khovd Summer Time") = 8
    TimeZoneDescs("Khovd Standard Time") = 7
    TimeZoneDescs("Indochina Time") = 7
    TimeZoneDescs("International Day Line West time zone") = -12
    TimeZoneDescs("Israel Daylight Time") = 3
    TimeZoneDescs("Indian Ocean Time") = 3
    TimeZoneDescs("Iran Daylight Time") = 4.5
    TimeZoneDescs("Irkutsk Time") = 8
    TimeZoneDescs("Iran Standard Time") = 3.5
    TimeZoneDescs("Indian Standard Time") = 5.5
    TimeZoneDescs("Irish Standard Time[5]") = 1
    TimeZoneDescs("Israel Standard Time") = 2
    TimeZoneDescs("Japan Standard Time") = 9
    TimeZoneDescs("Kaliningrad Time") = 2
    TimeZoneDescs("Kyrgyzstan Time") = 6
    TimeZoneDescs("Kosrae Time") = 11
    TimeZoneDescs("Krasnoyarsk Time") = 7
    TimeZoneDescs("Korea Standard Time") = 9
    TimeZoneDescs("Lord Howe Standard Time") = 10.5
    TimeZoneDescs("Lord Howe Summer Time") = 11
    TimeZoneDescs("Line Islands�Time") = 14
    TimeZoneDescs("Magadan Time") = 12
    TimeZoneDescs("Marquesas Islands Time") = -8.5
    TimeZoneDescs("Mawson Station Time") = 5
    TimeZoneDescs("Mountain Daylight Time�(North America)") = -6
    TimeZoneDescs("Middle European Time�Same zone as CET") = 1
    TimeZoneDescs("Middle European Summer Time�Same zone as CEST") = 2
    TimeZoneDescs("Marshall Islands Time") = 12
    TimeZoneDescs("Macquarie Island Station Time") = 11
    TimeZoneDescs("Marquesas Islands Time") = -8.5
    TimeZoneDescs("Myanmar Standard Time") = 6.5
    TimeZoneDescs("Moscow Time") = 3
    TimeZoneDescs("Malaysia Standard Time") = 8
    TimeZoneDescs("Mountain Standard Time�(North America)") = -7
    TimeZoneDescs("Mauritius Time") = 4
    TimeZoneDescs("Maldives Time") = 5
    TimeZoneDescs("Malaysia Time") = 8
    TimeZoneDescs("New Caledonia Time") = 11
    TimeZoneDescs("Newfoundland Daylight Time") = -1.5
    TimeZoneDescs("Norfolk Island Time") = 11
    TimeZoneDescs("Nepal Time") = 5.75
    TimeZoneDescs("Newfoundland Standard Time") = -2.5
    TimeZoneDescs("Newfoundland Time") = -2.5
    TimeZoneDescs("Niue Time") = -11
    TimeZoneDescs("New Zealand Daylight Time") = 13
    TimeZoneDescs("New Zealand Standard Time") = 12
    TimeZoneDescs("Omsk Time") = 6
    TimeZoneDescs("Oral Time") = 5
    TimeZoneDescs("Pacific Daylight Time�(North America)") = -7
    TimeZoneDescs("Peru Time") = -5
    TimeZoneDescs("Kamchatka Time") = 12
    TimeZoneDescs("Papua New Guinea Time") = 10
    TimeZoneDescs("Phoenix Island Time") = 13
    TimeZoneDescs("Philippine Time") = 8
    TimeZoneDescs("Pakistan Standard Time") = 5
    TimeZoneDescs("Saint Pierre and Miquelon Daylight Time") = -2
    TimeZoneDescs("Saint Pierre and Miquelon Standard Time") = -3
    TimeZoneDescs("Pohnpei Standard Time") = 11
    TimeZoneDescs("Pacific Standard Time�(North America)") = -8
    TimeZoneDescs("Philippine Standard Time") = 8
    TimeZoneDescs("Paraguay Summer Time[6]") = -3
    TimeZoneDescs("Paraguay Time[7]") = -4
    TimeZoneDescs("R�union Time") = 4
    TimeZoneDescs("Rothera Research Station Time") = -3
    TimeZoneDescs("Sakhalin Island Time") = 11
    TimeZoneDescs("Samara Time") = 4
    TimeZoneDescs("South African Standard Time") = 2
    TimeZoneDescs("Solomon Islands Time") = 11
    TimeZoneDescs("Seychelles Time") = 4
    TimeZoneDescs("Samoa Daylight Time") = -10
    TimeZoneDescs("Singapore Time") = 8
    TimeZoneDescs("Sri Lanka Standard Time") = 5.5
    TimeZoneDescs("Srednekolymsk Time") = 11
    TimeZoneDescs("Suriname Time") = -3
    TimeZoneDescs("Samoa Standard Time") = -11
    TimeZoneDescs("Singapore Standard Time") = 8
    TimeZoneDescs("Showa Station Time") = 3
    TimeZoneDescs("Tahiti Time") = -10
    TimeZoneDescs("Thailand Standard Time") = 7
    TimeZoneDescs("Indian/Kerguelen") = 5
    TimeZoneDescs("Tajikistan Time") = 5
    TimeZoneDescs("Tokelau Time") = 13
    TimeZoneDescs("Timor Leste Time") = 9
    TimeZoneDescs("Turkmenistan Time") = 5
    TimeZoneDescs("Turkey Time") = 3
    TimeZoneDescs("Tonga Time") = 13
    TimeZoneDescs("Tuvalu Time") = 12
    TimeZoneDescs("Ulaanbaatar Summer Time") = 9
    TimeZoneDescs("Ulaanbaatar Standard Time") = 8
    TimeZoneDescs("Coordinated Universal Time") = 0
    TimeZoneDescs("Uruguay Summer Time") = -2
    TimeZoneDescs("Uruguay Standard Time") = -3
    TimeZoneDescs("Uzbekistan Time") = 5
    TimeZoneDescs("Venezuelan Standard Time") = -4
    TimeZoneDescs("Vladivostok Time") = 10
    TimeZoneDescs("Volgograd Time") = 4
    TimeZoneDescs("Vostok Station Time") = 6
    TimeZoneDescs("Vanuatu�Time") = 11
    TimeZoneDescs("Wake Island Time") = 12
    TimeZoneDescs("West Africa Summer Time") = 2
    TimeZoneDescs("West Africa Time") = 1
    TimeZoneDescs("Western European Summer Time") = 1
    TimeZoneDescs("Western European Time") = 0
    TimeZoneDescs("Western Indonesian Time") = 7
    TimeZoneDescs("Western Standard Time") = 8
    TimeZoneDescs("Yakutsk Time") = 9
    TimeZoneDescs("Yekaterinburg Time") = 5
  end if
End Sub

'?stdDate.RegexMatch("asdf","(?<a>.)(.)((?<b>.)(.))")
Function RegexMatch(ByVal haystack As String, ByVal pattern As String, Optional ByVal options As String) As Object
  'Cache regexes for optimisation
  Static CachedRegex As Object
  Static CachedNames As Object
  If CachedRegex Is Nothing Then Set CachedRegex = CreateObject("Scripting.Dictionary")
  If CachedNames Is Nothing Then Set CachedNames = CreateObject("Scripting.Dictionary")
  
  'Named regexp used to detect capturing groups and named capturing groups
  Static NamedRegexp As Object
  If NamedRegexp Is Nothing Then
    Set NamedRegexp = CreateObject("VBScript.RegExp")
    NamedRegexp.pattern = "\((?:\?\<(.*?)\>)?"
    NamedRegexp.Global = True
  End If
  
  Static FreeSpace As Object
  If FreeSpace Is Nothing Then
    Set FreeSpace = CreateObject("VBScript.RegExp")
    FreeSpace.pattern = "\s+"
    FreeSpace.Global = True
  End If
  
  'If cached pattern doesn't exist, create it
  If Not CachedRegex(pattern) Then
    
    
    'Create names/capture group object
    Dim testPattern As String, oNames As Object
    testPattern = pattern
    testPattern = Replace(testPattern, "\\", "asdasd")
    testPattern = Replace(testPattern, "\(", "asdasd")
    
    'Store names for optimisation
    Set CachedNames(options & ")" & pattern) = NamedRegexp.Execute(testPattern)
    
    'Create new VBA valid pattern
    Dim newPattern As String
    newPattern = NamedRegexp.Replace(pattern, "(")
    
    'Create regexp from new pattern
    Dim oRegexp As Object
    Set oRegexp = CreateObject("VBScript.RegExp")
    
    'Set regex options
    Dim i As Integer
    For i = 1 To Len(flags)
        Select Case Mid(flags, i, 1)
            Case "i"
                oRegexp.ignoreCase = True
            Case "g"
                oRegexp.Global = True
            Case "x"
                newPattern = FreeSpace.Replace(newPattern, "(?:)")
            Case "s"
                newPattern = Replace(newPattern, "\\", "973ed556-6a75-45d6-b0c2-8c2d0e2431c9")
                newPattern = Replace(newPattern, "\.", "7ae2088d-1b1d-4ee1-8e38-956db49c12a2")
                newPattern = Replace(newPattern, ".", "(?:.|\s)")
                newPattern = Replace(newPattern, "7ae2088d-1b1d-4ee1-8e38-956db49c12a2", "\.")
                newPattern = Replace(newPattern, "973ed556-6a75-45d6-b0c2-8c2d0e2431c9", "\\")
            Case "m"
                oRegexp.MultiLine = True
        End Select
    Next
    
    'Set pattern
    oRegexp.pattern = newPattern
    
    'Store regex for optimisation
    
    Set CachedRegex(options & ")" & pattern) = oRegexp
  End If
  
  'Get matches object
  Dim oMatches As Object
  Set oMatches = CachedRegex(options & ")" & pattern).Execute(haystack)
  
  'Get names object
  Dim CName As Object
  Set CName = CachedNames(options & ")" & pattern)
  
  'Create dictionary to return
  Dim oRet As Object
  Set oRet = CreateObject("Scripting.Dictionary")
  
  'Fill dictionary with names and indexes
  '0 = Whole match
  '1,2,3,... = Submatch 1,2,3,...
  '"Count" stores the count of matches
  '"<<NAME>>" stores the match of a specified name
  For i = 1 To CName.Count
    oRet(i) = oMatches(0).Submatches(i - 1)
    If Not IsEmpty(CName(i - 1).Submatches(0)) Then oRet(CName(i - 1).Submatches(0)) = oMatches(0).Submatches(i - 1)
  Next i
  oRet(0) = oMatches(0)
  oRet("Count") = CName.Count
  Set RegexMatch = oRet
End Function

Friend Sub initialize()
  initialised = True
End Sub

'Used for unix time
Private Function dmod(a As Double, n As Double)
  dmod = a - (n * Int(a / n))
End Function

'Light weight function. Assumes format dd.mm.yyyy where . is any character
Private Function dd_mm_yyyy_2_value(str As String)
  Dim day, month, year As Integer
  day = CInt(Mid(Value, 1, 2))
  month = CInt(Mid(Value, 4, 2))
  year = CInt(Mid(Value, 7, 4))
  
  dd_mm_yyyy_2_value = CDbl(DateSerial(year, month, day))
End Function

'Get the GeoCode from the Windows API
'http://vbcity.com/forums/t/118955.aspx
Private Function getGeoCode() As String
  'See here for a better implementation:
  'http://vbcity.com/forums/t/118955.aspx
  
  'Assume UK
  getGeoCode = "UK"
End Function

