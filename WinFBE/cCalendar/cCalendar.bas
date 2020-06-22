' ========================================================================================
' WinFBE - FreeBASIC Editor (Windows 32/64 bit)
' Visual Designer auto generated project
' ========================================================================================

' Main application entry point.
' Place any additional global variables or #include files here.

' For your convenience, below are some of the most commonly used WinFBX library
' include files. Uncomment the files that you wish to use in the project or add
' additional ones. Refer to the WinFBX Framework Help documentation for information
' on how to use the various functions.

' #Include Once "Afx\AfxFile.inc"
' #Include Once "Afx\AfxStr.inc"
' #Include Once "Afx\AfxTime.inc"
' #Include Once "Afx\CIniFile.inc"
' #Include Once "Afx\CMoney.inc"
' #Include Once "Afx\CPrint.inc"
#Include Once "Afx\AfxTime.inc"
#Include Once "crInc\cCalendar\cCalendar.bi"
#Include Once "string.bi"

Dim shared oCalendar       as cCalendar
Dim Shared nSerial         as LongInt
Dim Shared uGregIn         as GREGORIAN_DATE
dim Shared uLocale         as LOCATION_LOCALE
dim Shared sDaylightName   as string
dim Shared sStandardName   as string
dim Shared sLocaleName     as string
Dim Shared arDaylight()    as DATE_CALCULATION
Dim Shared arUSAFederal()  as DATE_CALCULATION
dim shared nDaylightTime   as long = 0
dim Shared nStandardTime   as long = 0

' US Federal Holidays

    ReDim arUSAFederal(0 To 10)
    
' Holidays that fall on a Saturday are observed on Fri, Sunday on Monday

    arUSAFederal(0).Name = "New Year's Day"
    arUSAFederal(0).Month = cCalendarClass.JANUARY
    arUSAFederal(0).Day = 1
    arUSAFederal(0).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(0).Rule = cCalendarClass.NO_RULES
    arUSAFederal(0).Weekday = cCalendarClass.ALL_WEEKDAYS
    arUSAFederal(0).SaturdayRule = cCalendarClass.SATURDAY_OBSERVED_ON_FRIDAY
    arUSAFederal(0).SundayRule = cCalendarClass.SUNDAY_OBSERVED_ON_MONDAY
    arUSAFederal(0).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(0).NonBusinessDate = True

    arUSAFederal(1).Name = "Martin Luther King Day"    
    arUSAFederal(1).Month = cCalendarClass.JANUARY
    arUSAFederal(1).Day = 0
    arUSAFederal(1).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(1).Rule = cCalendarClass.THIRD_WEEK
    arUSAFederal(1).Weekday = cCalendarClass.MONDAY
    arUSAFederal(1).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
    arUSAFederal(1).SundayRule = cCalendarClass.NO_SUNDAY_RULE
    arUSAFederal(1).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(1).NonBusinessDate = True
    
' This holiday Is designated as "Washington’s Birthday" in section 6103(a) of title 5 of the United States Code,
' which is the law that specifies holidays for Federal employees.

    arUSAFederal(2).Name = "Washington's Birthday"    
    arUSAFederal(2).Month = cCalendarClass.FEBRUARY
    arUSAFederal(2).Day = 0
    arUSAFederal(2).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(2).Rule = cCalendarClass.THIRD_WEEK
    arUSAFederal(2).Weekday = cCalendarClass.MONDAY
    arUSAFederal(2).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
    arUSAFederal(2).SundayRule = cCalendarClass.NO_SUNDAY_RULE
    arUSAFederal(2).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(2).NonBusinessDate = True

    arUSAFederal(3).Name = "Memorial Day"    
    arUSAFederal(3).Month = cCalendarClass.MAY
    arUSAFederal(3).Day = 0
    arUSAFederal(3).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(3).Rule = cCalendarClass.LAST_WEEK
    arUSAFederal(3).Weekday = cCalendarClass.MONDAY
    arUSAFederal(3).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
    arUSAFederal(3).SundayRule = cCalendarClass.NO_SUNDAY_RULE
    arUSAFederal(3).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(3).NonBusinessDate = True

    arUSAFederal(4).Name = "Independence Day"    
    arUSAFederal(4).Month = cCalendarClass.JULY
    arUSAFederal(4).Day = 4
    arUSAFederal(4).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(4).Rule = cCalendarClass.NO_RULES
    arUSAFederal(4).Weekday = cCalendarClass.ALL_WEEKDAYS
    arUSAFederal(4).SaturdayRule = cCalendarClass.SATURDAY_OBSERVED_ON_FRIDAY
    arUSAFederal(4).SundayRule = cCalendarClass.SUNDAY_OBSERVED_ON_MONDAY
    arUSAFederal(4).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(4).NonBusinessDate = True

    arUSAFederal(5).Name = "Labor Day"    
    arUSAFederal(5).Month = cCalendarClass.SEPTEMBER
    arUSAFederal(5).Day = 0
    arUSAFederal(5).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(5).Rule = cCalendarClass.FIRST_WEEK
    arUSAFederal(5).Weekday = cCalendarClass.MONDAY
    arUSAFederal(5).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
    arUSAFederal(5).SundayRule = cCalendarClass.NO_SUNDAY_RULE
    arUSAFederal(5).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(5).NonBusinessDate = True

    arUSAFederal(6).Name = "Columbus Day"    
    arUSAFederal(6).Month = cCalendarClass.OCTOBER
    arUSAFederal(6).Day = 0
    arUSAFederal(6).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(6).Rule = cCalendarClass.SECOND_WEEK
    arUSAFederal(6).Weekday = cCalendarClass.MONDAY
    arUSAFederal(6).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
    arUSAFederal(6).SundayRule = cCalendarClass.NO_SUNDAY_RULE
    arUSAFederal(6).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(6).NonBusinessDate = True

    arUSAFederal(7).Name = "Veteran's Day"    
    arUSAFederal(7).Month = cCalendarClass.NOVEMBER
    arUSAFederal(7).Day = 11
    arUSAFederal(7).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(7).Rule = cCalendarClass.NO_RULES
    arUSAFederal(7).Weekday = cCalendarClass.ALL_WEEKDAYS
    arUSAFederal(7).SaturdayRule = cCalendarClass.SATURDAY_OBSERVED_ON_FRIDAY
    arUSAFederal(7).SundayRule = cCalendarClass.SUNDAY_OBSERVED_ON_MONDAY
    arUSAFederal(7).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(7).NonBusinessDate = True

    arUSAFederal(8).Name = "Thanksgiving Day"    
    arUSAFederal(8).Month = cCalendarClass.NOVEMBER
    arUSAFederal(8).Day = 0
    arUSAFederal(8).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(8).Rule = cCalendarClass.FOURTH_WEEK
    arUSAFederal(8).Weekday = cCalendarClass.THURSDAY
    arUSAFederal(8).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
    arUSAFederal(8).SundayRule = cCalendarClass.NO_SUNDAY_RULE
    arUSAFederal(8).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(8).NonBusinessDate = True

    arUSAFederal(9).Name = "Christmas Day"    
    arUSAFederal(9).Month = cCalendarClass.DECEMBER
    arUSAFederal(9).Day = 25
    arUSAFederal(9).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(9).Rule = cCalendarClass.NO_RULES
    arUSAFederal(9).Weekday = cCalendarClass.ALL_WEEKDAYS
    arUSAFederal(9).SaturdayRule = cCalendarClass.SATURDAY_OBSERVED_ON_FRIDAY
    arUSAFederal(9).SundayRule = cCalendarClass.SUNDAY_OBSERVED_ON_MONDAY
    arUSAFederal(9).YearRule = cCalendarClass.ALL_YEARS
    arUSAFederal(9).NonBusinessDate = True

    arUSAFederal(10).Name = "Inauguration Day"    
    arUSAFederal(10).Month = cCalendarClass.JANUARY
    arUSAFederal(10).Day = 20
    arUSAFederal(10).RuleClass = cCalendarClass.GREGORIAN_RULES
    arUSAFederal(10).Rule = cCalendarClass.NO_RULES
    arUSAFederal(10).Weekday = cCalendarClass.ALL_WEEKDAYS
    arUSAFederal(10).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
    arUSAFederal(10).SundayRule = cCalendarClass.SUNDAY_OBSERVED_ON_MONDAY
    arUSAFederal(10).YearRule = cCalendarClass.YEARS_AFTER_LEAP_ONLY
    arUSAFederal(10).NonBusinessDate = False

Redim arDaylight(0 to 1)

arDaylight(0).Name = "Daylight Savings Begins"    
arDaylight(0).Day = 0
arDaylight(0).Year = 0
arDaylight(0).RuleClass = cCalendarClass.GREGORIAN_RULES
arDaylight(0).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
arDaylight(0).SundayRule = cCalendarClass.NO_SUNDAY_RULE
arDaylight(0).YearRule = cCalendarClass.ALL_YEARS
arDaylight(0).NonBusinessDate = False
    
arDaylight(1).Name = "Daylight Savings Ends"    
arDaylight(1).Month = AfxTimeZoneStandardMonth()
arDaylight(1).Day = 0
arDaylight(1).Year = 0
arDaylight(1).RuleClass = cCalendarClass.GREGORIAN_RULES
arDaylight(1).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
arDaylight(1).SundayRule = cCalendarClass.NO_SUNDAY_RULE
arDaylight(1).YearRule = cCalendarClass.ALL_YEARS
arDaylight(1).NonBusinessDate = False

if arDaylight(1).Month <> 0 then
    uLocale.bApplyDaylightSavings = true
    arDaylight(0).Month = AfxTimeZoneDaylightMonth()
    arDaylight(0).Rule = AfxTimeZoneDaylightDay()
    arDaylight(0).Weekday = AfxTimeZoneDaylightDayOfWeek()
    nDaylightTime = AfxTimeZoneDaylightHour() * cCalendarClass.ONE_HOUR _
                  + AfxTimeZoneDaylightMinute() * cCalendarClass.ONE_MINUTE
    arDaylight(1).Rule = AfxTimeZoneStandardDay()
    arDaylight(1).Weekday = AfxTimeZoneStandardDayOfWeek()
    nStandardTime = AfxTimeZoneStandardHour() * cCalendarClass.ONE_HOUR _
                  + AfxTimeZoneStandardMinute() * cCalendarClass.ONE_MINUTE
else
    uLocale.bApplyDaylightSavings = false
end if

sDaylightName = AfxTimeZoneDaylightName()
sStandardName = AfxTimeZoneStandardName()
uLocale.Zone = AfxTimeZoneBias()
uLocale.DaylightSavingsMinutes = AfxTimeZoneDaylightBias()

' Assign Locale Name and Latitude/Longitude based on Zone

select case uLocale.Zone
    case -60
        sLocaleName = "Charles de Gaulle International Airport, Paris, Île-de-France, France"
        uLocale.Latitude = 49.012798
        uLocale.Longitude = 2.55
        uLocale.Elevation = 119
    case -120
        sLocaleName = "Cairo International Airport, Cairo, Cairo Governorate, Egypt"
        uLocale.Latitude = 30.121901
        uLocale.Longitude = 31.4056
        uLocale.Elevation = 116
    case -180
        sLocaleName = "Sheremetyevo International Airport, Moscow, Moscow Oblast, Russia"
        uLocale.Latitude = 55.972599
        uLocale.Longitude = 37.4146
        uLocale.Elevation = 190
    case -210
        sLocaleName = "Imam Khomeini International Airport, Tehran, Tehran Province, Iran"
        uLocale.Latitude = 35.4161
        uLocale.Longitude = 51.152199
        uLocale.Elevation = 1007
    case -240
        sLocaleName = "Abu Dhabi International Airport, Abu Dhabi, Abu Dhabi Emirate, UAE"
        uLocale.Latitude = 24.433001
        uLocale.Longitude = 54.6511
        uLocale.Elevation = 27
    case -270
        sLocaleName = "Kabul International Airport, Kabul, Kabul Province, Afghanistan"
        uLocale.Latitude = 34.565899
        uLocale.Longitude = 69.212303
        uLocale.Elevation = 1791
    case -300
        sLocaleName = "New Islamabad International Airport, Islamabad, Punjab, Pakistan"
        uLocale.Latitude = 33.560714
        uLocale.Longitude = 72.851614
        uLocale.Elevation = 502
    case -330
        sLocaleName = "Indira Gandhi International Airport, Delhi, Delhi, India"
        uLocale.Latitude = 28.5665
        uLocale.Longitude = 77.103104
        uLocale.Elevation = 237
    case -345
        sLocaleName = "Tribhuvan International Airport, Kathmandu, Bagmati, Nepal"
        uLocale.Latitude = 27.6966
        uLocale.Longitude = 85.3591
        uLocale.Elevation = 1338 
    case -360
        sLocaleName = "Dhaka/Hazrat Shahjalal International, Dhaka, Dhaka Division, Bangladesh"
        uLocale.Latitude = 23.843347
        uLocale.Longitude = 90.397783
        uLocale.Elevation = 9
    case -390
        sLocaleName = "Yangon International Airport, Yangon(Rangoon), Yangon Division, Myanmar"
        uLocale.Latitude = 16.907301
        uLocale.Longitude = 96.133202
        uLocale.Elevation = 33
    case -420
        sLocaleName = "Don Mueang International Airport, Bangkok, Bangkok Province, Thailand"
        uLocale.Latitude = 13.9126
        uLocale.Longitude = 100.607002
        uLocale.Elevation = 3                                              
    case -480
        sLocaleName = "Guangzhou Baiyun International Airport, Guangdong Province China"
        uLocale.Latitude = 23.38831778
        uLocale.Longitude = 113.2992665496
        uLocale.Elevation = 15
    case -525
        sLocaleName = "Eucla, Western Australia, Australia"
        uLocale.Latitude = -31.67713
        uLocale.Longitude = 128.8893
        uLocale.Elevation = 93
    case -540
        sLocaleName = "Gimpo International Airport, Seoul, Seoul Teugbyeolsi, South Korea"
        uLocale.Latitude = 37.5583
        uLocale.Longitude = 126.791
        uLocale.Elevation = 18
    case -570
        sLocaleName = "Darwin International Airport, Darwin, Northern Territory, Australia"
        uLocale.Latitude = -12.4147
        uLocale.Longitude = 130.876999
        uLocale.Elevation = 31
    case -600
        sLocaleName = "Melbourne International Airport, Melbourne, Victoria, Australia"
        uLocale.Latitude = -37.673302
        uLocale.Longitude = 144.843002
        uLocale.Elevation = 132
    case -630
        sLocaleName = "Lord Howe Island Airport, Lord Howe Island, New South Wales, Australia"
        uLocale.Latitude = -31.5383
        uLocale.Longitude = 159.076996
        uLocale.Elevation = 2
    case -660
        sLocaleName = "Honiara International Airport, Honiara, Solomon Islands"
        uLocale.Latitude = -9.428
        uLocale.Longitude = 160.054993
        uLocale.Elevation = 9
    case -720
        sLocaleName = "Nadi International Airport, Nadi, Western, Fiji"
        uLocale.Latitude = -17.7554
        uLocale.Longitude = 177.442993
        uLocale.Elevation = 18
    case -765
        sLocaleName = "Tuuta Airport, Chatham Island, New Zealand"
        uLocale.Latitude = -43.810001
        uLocale.Longitude = -176.457001
        uLocale.Elevation = 13
    case -780
        sLocaleName = "Fua'amotu International Airport, Nuku'alofa, Tongatapu, Tonga"
        uLocale.Latitude = -21.241199
        uLocale.Longitude = -175.149994
        uLocale.Elevation = 38                                     
    case 0
        sLocaleName = "Heathrow Airport, Hounslow, Greater London, UK"
        uLocale.Latitude = 51.471439
        uLocale.Longitude = -0.456802
        uLocale.Elevation = 25    
    case 60
        sLocaleName = "Praia International Airport, Praia, Sotavento Islands, Cape Verde"
        uLocale.Latitude = 14.9245
        uLocale.Longitude = -23.4935
        uLocale.Elevation = 70    
    case 120
        sLocaleName = "Stanley, Falkland Islands"
        uLocale.Latitude = 51.70098
        uLocale.Longitude = -57.849186
        uLocale.Elevation = 37     
    case 180
        sLocaleName = "Halifax / Stanfield International Airport, Nova Scotia, Canada"
        uLocale.Latitude = 44.880798
        uLocale.Longitude = -63.508598
        uLocale.Elevation = 145    
    case 210
        sLocaleName = "Gander International Airport, Gander, Newfoundland, Canada"
        uLocale.Latitude = 48.936901
        uLocale.Longitude = -54.5681
        uLocale.Elevation = 151    
    case 240
        sLocaleName = "Luis Munoz Marin International Airport, San Juan, Puerto Rico, USA"
        uLocale.Latitude = 18.439399
        uLocale.Longitude = -66.002133
        uLocale.Elevation = 3    
    case 300
        sLocaleName = "John F Kennedy International Airport, New York, New York, USA"
        uLocale.Latitude = 40.6399278
        uLocale.Longitude = -73.7786925
        uLocale.Elevation = 4        
    case 360
        sLocaleName = "Des Moines International Airport, Des Moines, Iowa, USA"
        uLocale.Latitude = 41.5339722
        uLocale.Longitude = -93.6630833
        uLocale.Elevation = 292    
    case 420
        sLocaleName = "Salt Lake City International Airport, Salt Lake City, Utah, USA"
        uLocale.Latitude = 40.7883933
        uLocale.Longitude = -111.9777733
        uLocale.Elevation = 1290     
    case 480
        sLocaleName = "Spokane International Airport, Spokane, Washington, USA"
        uLocale.Latitude = 47.6190278
        uLocale.Longitude = -117.5352222
        uLocale.Elevation = 727       
    case 540
        sLocaleName = "Ted Stevens International Airport, Anchorage, Alaska, USA"
        uLocale.Latitude = 61.1740847
        uLocale.Longitude = -149.9981375
        uLocale.Elevation = 46
    case 570
        sLocaleName = "Faa'a International Airport, Papeete, French Polynesia, France"
        uLocale.Latitude = -17.553699
        uLocale.Longitude = -149.606995
        uLocale.Elevation = 2     
    case 600
        sLocaleName = "Daniel K Inouye International Airport, Honululu, Hawaii, USA"
        uLocale.Latitude = 21.3178275
        uLocale.Longitude = -157.9202627
        uLocale.Elevation = 4    
    case 660
        sLocaleName = "Pago Pago International Airport, American Samoa"
        uLocale.Latitude = -14.3316622
        uLocale.Longitude = -170.7115031
        uLocale.Elevation = 10    
    case 720
        sLocaleName = "Wake Island Airfield, USA Minor Outlying Islands"
        uLocale.Latitude = 19.282489
        uLocale.Longitude = 166.636661
        uLocale.Elevation = 7    
end select

uLocale.Zone = (uLocale.Zone / 60.0) * -1


Application.Run(frmMain)
