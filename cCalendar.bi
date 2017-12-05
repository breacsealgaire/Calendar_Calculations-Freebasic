' ########################################################################################
' File: cCalendar.bi
' Contents: Algorithms for various calendars of current and historical interest.
' Version: 1.01
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2016 Rick Kelly
' Credits - Calendrical Calculations Third Edition, Nachum Dershowitz and Edward M. Reingold
'           Astronomical Algorithms Second Edition, Jean Meeus
' Released into the public domain for private and public use without restriction
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#Pragma Once

Namespace cCalendarClass

' Constants

' Epoch definitions in Days

Private Const GREGORIAN_EPOCH              =   1                         ' January 1, 0001
Private Const JULIAN_EPOCH                 =   -1                        ' December 30, 0000
Private Const EGYPTIAN_EPOCH               =   -272787                   ' February 26, -747
Private Const ARMENIAN_EPOCH               =   201443                    ' July 11, 552
Private Const COPTIC_EPOCH                 =   103605                    ' August 29, 0284
Private Const ETHIOPIC_EPOCH               =   2796                      ' August 29, 8
Private Const ISLAMIC_EPOCH                =   227015                    ' July 19, 622
Private Const HEBREW_EPOCH                 =   -1373427                  ' October 7, -3761 (Julian)
Private Const HINDU_EPOCH                  =   -1132959                  ' January 23, -3101
Private Const CHINESE_EPOCH                =   -963099                   ' Feb 15 -2636
Private Const CHINESE_MONTH_NAME_EPOCH     =   57
Private Const PERSIAN_EPOCH                =   226896                    ' March 22, 622
Private Const BAHAI_EPOCH                  =   673222                    ' March 21, 1844
Private Const TIBETAN_EPOCH                =   -46410                    ' December 7, -127  
Private Const EXCEL_1900_EPOCH as LongInt  =   693596                    ' January 1, 1900
Private Const EXCEL_1904_EPOCH as LongInt  =   695056                    ' January 1, 1904
Private Const UNIX_EPOCH as LongInt        =   719163                    ' January 1, 1970
                                                                         
Private Const J2000 as Double                         =   730120.5       ' January 1, 2000 at noon
Private Const TehranLocale_Latitude as Double         =   35.69439  
Private Const TehranLocale_Longitude as Double        =   51.42151 
Private Const TehranLocale_Elevation as Long          =   1178
Private Const TehranLocale_Zone as Double             =   3.5
Private Const HinduLocaleZone as Double               =   5 + 461 / 9000
Private Const HinduLocaleElevation as Long            =   0
Private Const HinduLocaleLongitude as Double          =   75 + (46 / 60) + (6 / 3600)               ' Ujjain 75 deg, 46 min, 6 sec longitude
Private Const HinduSiderealYear as Double             =   (365 + 279457 / 1080000)
Private Const HinduSiderealMonth as Double            =   (27 + 4644439 / 14438334)
Private Const HinduSynodicMonth as Double             =   (29 + 7087771 / 13358334)
Private Const HinduCreation as Double                 =   (HINDU_EPOCH - 1955880000 * HinduSiderealYear)
Private Const HinduAnomalisticYear as Double          =   (1577917828000 / (4320000000 - 387))
Private Const HinduAnomalisticMonth as Double         =   (1577917828 / (57753336 - 488199))
Private Const HinduLocaleLatitude as Double           =   23 + (9 / 60)                             ' Ujjain 23 deg, 9 seconds latitude

' Meeus 'Astronomical Algorithms, 2nd ed of June 15, 2005

Private Const MeanSynodicMonth as Double              =   29.530588861                              ' Mean time from new moon to new moon
Private Const VisibleHorizon as Double                =   0.8413147543981382                        ' Half diameter of the sun (16 minutes +
                                                                                                    ' 34.478885263888294 minutes for refraction)
                                                                                                    ' expressed in decimal degrees. Since the moon
                                                                                                    ' and sun are very close to same angular size,
                                                                                                    ' refraction for both uses this value
 
' Time

Private Const ONE_DAY as LongInt    =   86400000                              ' Milliseconds in a day
Private Const ONE_HOUR as LongInt   =   3600000                               ' Milliseconds in a hour
Private Const ONE_MINUTE as LongInt =   60000                                 ' Milliseconds in a minute
Private Const ONE_SECOND as LongInt =   1000                                  ' Milliseconds in a second

' Gregorian

Private Const JANUARY           =   1
Private Const FEBRUARY          =   2
Private Const MARCH             =   3
Private Const APRIL             =   4
Private Const MAY               =   5
Private Const JUNE              =   6
Private Const JULY              =   7
Private Const AUGUST            =   8
Private Const SEPTEMBER         =   9
Private Const OCTOBER           =   10
Private Const NOVEMBER          =   11
Private Const DECEMBER          =   12

Private Const ALL_WEEKDAYS      =   -1
Private Const SUNDAY            =   0
Private Const MONDAY            =   1
Private Const TUESDAY           =   2
Private Const WEDNESDAY         =   3
Private Const THURSDAY          =   4
Private Const FRIDAY            =   5
Private Const SATURDAY          =   6

' Hebrew

Private Const NISAN             = 1
Private Const IYAR              = 2
Private Const SIVAN             = 3
Private Const TAMMUZ            = 4
Private Const AV                = 5
Private Const ELUL              = 6
Private Const TISHRI            = 7
Private Const MARHESHVAN        = 8
Private Const KISLEV            = 9
Private Const TEVET             = 10
Private Const SHEVAT            = 11
Private Const ADAR              = 12
Private Const ADARII            = 13

' Chinese

Private Const CHINESE           = 1
Private Const VIETNAMESE        = 2
Private Const KOREAN            = 3
Private Const JAPANESE          = 4

Private Const CHINESE_RAT       = 1
Private Const CHINESE_OX        = 2
Private Const CHINESE_TIGER     = 3
Private Const CHINESE_RABBIT    = 4
Private Const CHINESE_DRAGON    = 5
Private Const CHINESE_SNAKE     = 6
Private Const CHINESE_HORSE     = 7
Private Const CHINESE_GOAT      = 8
Private Const CHINESE_MONKEY    = 9
Private Const CHINESE_ROOSTER   = 10
Private Const CHINESE_DOG       = 11
Private Const CHINESE_PIG       = 12

Private Const CHINESE_WIDOW_YEAR           = 1
Private Const CHINESE_BLIND_YEAR           = 2
Private Const CHINESE_BRIGHT_YEAR          = 3
Private Const CHINESE_DOUBLE_BRIGHT_YEAR   = 4


' Hindu

Private Const HINDU_SOLAR_ERA   = 3179
Private Const HINDU_LUNAR_ERA   = 3044

' Solar

Private Const VAISAKHA          = 1
Private Const JYAISHTHA         = 2
Private Const ASHADHA           = 3
Private Const SRAVANA           = 4
Private Const BRADRAPADA        = 5
Private Const ASVINA            = 6
Private Const KARTIKA           = 7
Private Const MARGASIRSHA       = 8
Private Const PAUSHA            = 9
Private Const MAGHA             = 10
Private Const PHALGUNA          = 11
Private Const CHAITRA           = 12
 
' Roman

Private Const KALENDS           =   1
Private Const NONES             =   2
Private Const IDES              =   3

' Coptic
                        
Private Const THOOUT            =   1
Private Const PAOPE             =   2
Private Const ATHOR             =   3
Private Const KOIAK             =   4
Private Const TOBE              =   5
Private Const MESHIR            =   6
Private Const PAREMOTEP         =   7
Private Const PARMOUTE          =   8
Private Const PASHONS           =   9
Private Const PAONE             =   10
Private Const EPEP              =   11
Private Const MESORE            =   12
Private Const EPAGOMENE         =   13

' Ethiopic

Private Const MASKARAM          =   1
Private Const TEQEMT            =   2
Private Const HEDAR             =   3
Private Const TAKHSAS           =   4
Private Const TER               =   5
Private Const YAKATIT           =   6
Private Const MAGABIT           =   7
Private Const MIYAZYA           =   8
Private Const GENBOT            =   9
Private Const SANE              =   10
Private Const HAMLE             =   11
Private Const NAHASE            =   12
Private Const PAGUEMEN          =   13

' Persian

Private Const FARVARDIN         =   1
Private Const ORDIBEHESHT       =   2
Private Const XORDAD            =   3
Private Const TIR               =   4
Private Const MORDAD            =   5
Private Const SHAHRIVAR         =   6
Private Const MEHR              =   7
Private Const ABAN              =   8
Private Const AZAR              =   9
Private Const DEY               =   10
Private Const BAHMAN            =   11
Private Const ESFAND            =   12

' Islamic

Private Const MUHARRAM          =  1
Private Const SAFAR             =  2
Private Const RABIALAWWAL       =  3
Private Const RABIALAHIR        =  4
Private Const JUMADAALULA       =  5
Private Const JUMADAALAHIRA     =  6
Private Const RAJAB             =  7
Private Const SHABAN            =  8
Private Const RAMADAN           =  9
Private Const SHAWWAL           =  10
Private Const DHUALQADA         =  11
Private Const DHUALHIJJA        =  12

' Bahai

Private Const BAHA              =   1
Private Const JALAL             =   2
Private Const JAMAL             =   3
Private Const AZAMAT            =   4
Private Const NUR               =   5
Private Const RAHMAT            =   6
Private Const KALIMAT           =   7
Private Const KAMAL             =   8
Private Const ASMA              =   9
Private Const IZZAT             =   10
Private Const MASHIYYAT         =   11
Private Const ILM               =   12
Private Const QUDRAT            =   13
Private Const QAWL              =   14
Private Const MASAIL            =   15
Private Const SHARAF            =   16
Private Const SULTAN            =   17
Private Const MULK              =   18
Private Const AYYAMIHA          =   0
Private Const ALA               =   19

' Tibetan

Private Const DBO               =   1
Private Const NAGPA             =   2
Private Const SAGA              =   3
Private Const SNRON             =   4
Private Const CHUSTOD           =   5
Private Const GROBZHIN          =   6
Private Const KHRUMS            =   7
Private Const THASKAR           =   8
Private Const SMINDRUG          =   9
Private Const MGO               =   10
Private Const RGYAL             =   11
Private Const MCHU              =   12

' Astronomical definitions

Private Const PI as Double       =   3.14159265358979323846
Private Const SPRING             =   0
Private Const SUMMER             =   90
Private Const AUTUMN             =   180
Private Const WINTER             =   270

Private Const SUNRISE_SUNSET_TIME           = 0
Private Const CIVIL_TWILIGHT_TIME           = 6 
Private Const NAUTICAL_TWIGHTLIGHT_TIME     = 12
Private Const ASTRONOMICAL_TWILIGHT_TIME    = 18

Private Const NEWMOON                   =   0
Private Const FIRSTQUARTERMOON          =   90
Private Const FULLMOON                  =   180
Private Const LASTQUARTERMOON           =   270
Private Const GEOCENTRIC as BOOLEAN     = True
Private Const TOPOCENTRIC as BOOLEAN    = False
Private Const MOONRISE as BOOLEAN       = True
Private Const MOONSET as BOOLEAN        = False   
Private Const MORNING as BOOLEAN        = True
Private Const EVENING as BOOLEAN        = False

' Rule options - used when Weekday is Sunday through Saturday (0-6)

Private Const NO_RULES                       = 0
Private Const FIRST_WEEK                     = 1
Private Const SECOND_WEEK                    = 2
Private Const THIRD_WEEK                     = 3
Private Const FOURTH_WEEK                    = 4
Private Const LAST_WEEK                      = 5
Private Const LAST_FULL_WEEK                 = 6
Private Const BEFORE                         = 7
Private Const ON_OR_BEFORE                   = 8
Private Const AFTER                          = 9
Private Const ON_OR_AFTER                    = 10
Private Const NEAREST                        = 11

' Rules for Rule Class = CHRISTIANEASTER_RULES

Private Const CHRISTIAN_EASTER_EASTER            = 12
Private Const CHRISTIAN_EASTER_GOODFRIDAY        = 13
Private Const CHRISTIAN_EASTER_MAUNDYTHURSDAY    = 14
Private Const CHRISTIAN_EASTER_PALMSUNDAY        = 15
Private Const CHRISTIAN_EASTER_PASSIONSUNDAY     = 15           ' Palm Sunday and Passion Sunday are observed on the same day currently
Private Const CHRISTIAN_EASTER_MOTHERINGSUNDAY   = 17           ' Observed in some parts of Europe
Private Const CHRISTIAN_EASTER_ASHWEDNESDAY      = 18
Private Const CHRISTIAN_EASTER_MARDIGRAS         = 19
Private Const CHRISTIAN_EASTER_SHROVEMONDAY      = 20
Private Const CHRISTIAN_EASTER_SHROVESUNDAY      = 21
Private Const CHRISTIAN_EASTER_SEXAGESIMASUNDAY  = 22
Private Const CHRISTIAN_EASTER_EASTERMONDAY      = 23
Private Const CHRISTIAN_EASTER_ROGATIONSUNDAY    = 24
Private Const CHRISTIAN_EASTER_ASCENSION         = 25
Private Const CHRISTIAN_EASTER_PENTECOST         = 26
Private Const CHRISTIAN_EASTER_TRINITYSUNDAY     = 27
Private Const CHRISTIAN_EASTER_CORPUSCHRISTI     = 28

' Rules for Rule Class = ORTHODOXEASTER_RULES

Private Const ORTHODOX_EASTER_NEWYEAR            = 29
Private Const ORTHODOX_EASTER_EASTER             = 30
Private Const ORTHODOX_EASTER_GOODFRIDAY         = 31
Private Const ORTHODOX_EASTER_PALMSUNDAY         = 32
Private Const ORTHODOX_EASTER_FORGIVENESSSUNDAY  = 33
Private Const ORTHODOX_EASTER_GREATLENT          = 34
Private Const ORTHODOX_EASTER_FEASTOFASCENSION   = 35
Private Const ORTHODOX_EASTER_PENTECOST          = 36
Private Const ORTHODOX_EASTER_APOSTLESFAST       = 37
Private Const ORTHODOX_EASTER_ALLSAINTSSUNDAY    = 38

' Rules for Rule Class = HEBREW_RULES

Private Const HEBREW_HESHVAN30                   = 39
Private Const HEBREW_KISLEV30                    = 40

' Rules for Rule Class = CHINESE_RULES (Vietnamese/Korean/Japanese included)

Private Const CHINESE_WINTERSOLSTICE             = 41
Private Const CHINESE_QINGMING                   = 42

' Rules for Rule Class = ISLAMIC_RULES

Private Const ISLAMIC_QUDS_DAY                   = 43

' Rules for Rule Class = HINDU_SOLAR_RULES

Private Const HINDU_SOLAR_MESHA_SAMKRANTI        = 1

' Rules for Rule Class = BAHAI_RULES

Private Const BIRTH_OF_BAB                   = 1
Private Const BIRTH_OF_BAHAULLAH             = 2
Private Const NAW_RUZ                        = 3

' Rules for Rule Class = TIBETAN_RULES

Private Const LOSAR                          = 1

' Thursday options (when date falls on a Thursday)

Private Const NO_THURSDAY_RULE               = 0
Private Const THURSDAY_OBSERVED_ON_WEDNESDAY = 1

' Friday options (when date falls on a Friday)

Private Const NO_FRIDAY_RULE                 = 0
Private Const FRIDAY_OBSERVED_ON_THURSDAY    = 1
Private Const FRIDAY_OBSERVED_ON_WEDNESDAY   = 2

' Saturday options (when date falls on a Saturday)

Private Const NO_SATURDAY_RULE               = 0
Private Const SATURDAY_OBSERVED_ON_FRIDAY    = 1
Private Const SATURDAY_OBSERVED_ON_MONDAY    = 2
Private Const SATURDAY_OBSERVED_ON_SUNDAY    = 3
Private Const SATURDAY_OBSERVED_ON_THURSDAY  = 4

' Sunday options (when date falls on a Sunday)

Private Const NO_SUNDAY_RULE                 = 0
Private Const SUNDAY_OBSERVED_ON_MONDAY      = 1

' Monday options (when date falls on a Monday)

Private Const NO_MONDAY_RULE                 = 0
Private Const MONDAY_OBSERVED_ON_TUESDAY     = 1


' Year options (return of False if Year fails test)

Private Const ALL_YEARS                      = 0
Private Const ODD_YEARS_ONLY                 = 1
Private Const EVEN_YEARS_ONLY                = 2
Private Const YEARS_AFTER_LEAP_ONLY          = 3        ' The Julian Leap Year calculation is used (evenly divisible by 4)
Private Const LEAP_YEARS_ONLY                = 4        ' The Julian Leap Year calculation is used (evenly divisible by 4)

' Holiday Rule Class

Private Const GREGORIAN_RULES                = 0
Private Const CHRISTIANEASTER_RULES          = 1
Private Const ORTHODOXEASTER_RULES           = 2
Private Const HEBREW_RULES                   = 3
Private Const CHINESE_RULES                  = 4
Private Const KOREAN_RULES                   = 5
Private Const JAPANESE_RULES                 = 6
Private Const VIETNAMESE_RULES               = 7
Private Const PERSIAN_RULES                  = 8
Private Const HINDU_SOLAR_RULES              = 9
Private Const HINDU_LUNAR_RULES              = 10
Private Const ISLAMIC_RULES                  = 11
Private Const COPTIC_RULES                   = 12
Private Const ETHIOPIC_RULES                 = 13
Private Const BAHAI_RULES                    = 14
Private Const TIBETAN_RULES                  = 15

' Date Validation

Private Const VALID_DATE                     = 0
Private Const INVALID_MONTH                  = 1
Private Const INVALID_DAY                    = 2
Private Const INVALID_YEAR                   = 3
Private Const INVALID_WEEK                   = 4
Private Const INVALID_CYCLE                  = 5
Private Const INVALID_LEAP_MONTH             = 6
Private Const INVALID_LEAP_DAY               = 7
Private Const INVALID_MAJOR                  = 8
Private Const INVALID_EVENT                  = 9
Private Const INVALID_COUNT                  = 10
Private Const INVALID_LEAP_YEAR              = 11


End Namespace

' Defined date types

Private Type GREGORIAN_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    Weekday             as Short
    LeapYear            as BOOLEAN
    
End Type

Private Type ISO_DATE

    Week                as Short
    Day                 as Short
    Year                as Long
    LongYear            as BOOLEAN
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
                  
End Type

Private Type ISLAMIC_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type

Private Type HEBREW_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    SabbaticalYear      as BOOLEAN
    
End Type

Private Type PERSIAN_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type

Private Type HINDU_SOLAR_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type

Private Type HINDU_LUNAR_DATE

    Month               as Short
    LeapMonth           as BOOLEAN
    Day                 as Short
    LeapDay             as BOOLEAN
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type

Private Type JULIAN_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    Weekday             as Short
    LeapYear            as BOOLEAN
    
End Type

Private Type ROMAN_DATE

    Month               as Short
    Year                as Long
    Event               as Short
    Count               as Short
    Leap                as BOOLEAN
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
                         
End Type

Private Type EGYPTIAN_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type

Private Type ARMENIAN_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type

Private Type ETHIOPIC_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type

Private Type COPTIC_DATE

    Month               as Short
    Day                 as Short
    Year                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type

Private Type CHINESE_DATE

    Cycle               as Short
    Year                as Long
    YearAnimal          as Short
    YearAugury          as Short
    Month               as Short
    MonthAnimal         as Short
    LeapMonth           as BOOLEAN
    Day                 as Short
    Country             as Short                    ' Must be set to CHINESE,VIETNAMESE,KOREAN,JAPANESE
                                                    ' when converting to or from Chinese. Invalid country
                                                    ' will default to CHINESE 
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
                                
End Type

Private Type BAHAI_DATE

    Major               as Short
    Cycle               as Short
    Month               as Short
    Day                 as Short
    Year                as Short
    Era                 as Short
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type

Private Type TIBETAN_DATE

    Month               as Short
    LeapMonth           as BOOLEAN
    Day                 as Short
    LeapDay             as BOOLEAN
    Year                as Long 
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
                                
End Type

' UDT's

Private Type TIME_DURATION

    Days                as Long
    Hour                as Short
    Minute              as Short
    Second              as Short
    Millisecond         as Short
    
End Type  

Private Type HISTORY_MONTHS

' 0 = Current Month, 1=Jan...12=Dec

    Month(0 To 12)          as Double

End Type

' Date Calculations and Validations

Private Type DATE_CALCULATION

    Name                as String
    Month               as Short
    Day                 as Short
    Year                as Long          ' Always a Gregorian Year
    RuleClass           as Short
    Rule                as Short
    Weekday             as Short
    ThursdayRule        as Short
    FridayRule          as Short
    SaturdayRule        as Short
    SundayRule          as Short
    MondayRule          as Short
    YearRule            as Short
    WeekRule            as Short
    WeekRuleWeekday     as Short         ' First day of a week - value range Sun-Sat
                                         ' Used in business date calculations
    Observed            as BOOLEAN       ' TRUE if date occurs in nYear, FALSE otherwise
    ObservedDays1       as LongInt
    ObservedDaysBegin1  as Long          ' First day of Observed date week in DAYS format
    ObservedDays2       as LongInt
    ObservedDaysBegin2  as Long          ' First day of Observed date week in DAYS format
    NonBusinessDate     as BOOLEAN       ' TRUE if date is skipped during business date calculations
    MaxNonBusinessDates as BOOLEAN       ' TRUE if capacity reached for saving dates

End Type

Private Type LOCATION_LOCALE
    
     Latitude                       as Double   
     Longitude                      as Double   
     Zone                           as Double             ' In hours, set to < 0 when Longitude is < 0
     bApplyDaylightSavings          as BOOLEAN
     DaylightSavingsBegins          as LongInt
     DaylightSavingsEnds            as LongInt
     DaylightSavingsMinutes         as Short
     bDaylightLightSavingsActive    as BOOLEAN   
     Elevation                      as Long               ' In meters
        
End Type   

Private Type LUNAR_PHASES
   
     LunarSerialTIme         as LongInt   
     Phase                   as Short              ' Equal to NEWMOON,FIRSTQUARTERMOON,FULLMOON,LASTQUARTERMOON
       
End Type

Private Type LUNAR_RISE_AND_SET 
 
     LunarSerialTime         as LongInt   
     RiseOrSet               as BOOLEAN            ' False=Moonset, True=Moonrise
     DaylightSavings         as BOOLEAN            ' True if Daylight Savings adjustment was made
      
End Type 

' The basic date and time format is a 64 bit LONGINT representing the number of
' milliseconds since January 1, 1 at midnight referred to as a serial date.

' A moment is a double precision value representing the days since January 1, 1
' with the fractional part representing a portion of one day.

' Calculations involving astronomical events use algorithms that are fairly precise
' within +- 2000 years or so. Outside that range, errata increase the farther from
' that range. Rise and Set times are +- 10 min or so from published values

' ########################################################################################
' cCalendar Class
' ########################################################################################

Type cCalendar Extends Object

   Private:
   
      arNonBusinessDays(0 To 999)   as Long
      iUBound                       as Long
      arBusinessWeekdays(0 To 6)    as BOOLEAN
      
' Date Calculation Support

      Declare Function cmDateCalculationObserved(ByVal nThursdayOption as Short, _
                                                 ByVal nFridayOption as Short, _
                                                 ByVal nSaturdayOption as Short, _
                                                 ByVal nSundayOption as Short, _
                                                 ByVal nMondayOption as Short, _
                                                 ByVal nYearsOption as Short, _
                                                 ByVal nDays as Long, _
                                                 ByVal nYear as Long, _
                                                 ByRef nObservedDays as LongInt) as BOOLEAN
      Declare Sub cmSaveDateCalculation(ByRef uCalc as DATE_CALCULATION)

' Gregorian support

      Declare Sub cmGregorianDateCalculation(ByVal nMonth as Short, _
                                             ByVal nDay as Short, _
                                             ByVal nYear as Long, _
                                             ByVal nWeekDay as Short, _
                                             ByVal nRule as Short, _
                                             ByRef nCalcDays as Long)
      Declare Function cmGregorianLeapYear (ByVal nYear as Long) as BOOLEAN
      Declare Function cmDaysFromGregorian(ByVal nMonth as Long, _
                                           ByVal nDay as Long, _
                                           ByVal nYear as Long) as Long
      Declare Sub cmGregorianFromDays(ByVal nDays as Long, _
                                      ByRef nMonth as Short, _
                                      ByRef nDay as Short, _
                                      ByRef nYear as Long)
      Declare Function cmGregorianWeekDay (ByVal nDays as Long) as Short
      Declare Function cmGregorianYearFromDays (ByVal nDays as Long) as Long
      Declare Sub cmGregorianYearRange(ByVal nYear as Long, _
                                       ByRef nYearStart as Long, _
                                       ByRef nYearEnd as Long)
      Declare Function cmGregorianNewYear(ByVal nYear as Long) as Long
      Declare Function cmGregorianYearEnd(ByVal nYear as Long) as Long
      Declare Function cmFirstWeekDay (ByVal nWeekDay as Short, _
                                       ByVal nDays as Long) as Long
      Declare Function cmSecondWeekDay (ByVal nWeekDay as Short, _
                                        ByVal nDays as Long) as Long
      Declare Function cmThirdWeekDay (ByVal nWeekDay as Short, _
                                       ByVal nDays as Long) as Long
      Declare Function cmFourthWeekDay (ByVal nWeekDay as Short, _
                                        ByVal nDays as Long) as Long
      Declare Function cmLastWeekDay (ByVal nWeekDay as Short, _
                                      ByVal nDays as Long) as Long
      Declare Function cmNthWeekDay (ByVal nNthDay as Short, _
                                     ByVal nWeekDay as Short, _
                                     ByVal nDays as Long) as Long
      Declare Function cmWeekDayBefore (ByVal nWeekDay as Short, _
                                        ByVal nDays as Long) as Long
      Declare Function cmWeekDayAfter (ByVal nWeekDay as Short, _
                                       ByVal nDays as Long) as Long
      Declare Function cmWeekDayNearest (ByVal nWeekday as Short, _
                                         ByVal nDays as Long) as Long
      Declare Function cmWeekDayOnOrAfter (ByVal nWeekDay as Short, _
                                           ByVal nDays as Long) as Long
      Declare Function cmWeekDayOnOrBefore (ByVal nWeekDay as Short, _
                                            ByVal nDays as Long) as Long
      Declare Function cmGregorianDateDifference(ByVal nStartMonth as Short, _
                                                 ByVal nStartDay as Short, _
                                                 ByVal nStartYear as Long, _
                                                 ByVal nEndMonth as Short, _
                                                 ByVal nEndDay as Short, _
                                                 ByVal nEndYear as Long) as Long
                                                 
' ISO Support

      Declare Function cmISOLongYear (ByVal nYear as Long) as BOOLEAN
      Declare Sub cmISOFromDays(ByVal nDays as Long, _
                                ByRef nWeek as Short, _
                                ByRef nDay as Short, _
                                ByRef nYear as Long)
      Declare Function cmDaysFromISO(ByVal nWeek as Short, _
                                     ByVal nDay as Short, _
                                     ByVal nYear as Long) as Long
                                                 
' Julian support

      Declare Sub cmJulianFromDays(ByVal nDays as Long, _
                                   ByRef nMonth as Short, _
                                   ByRef nDay as Short, _
                                   ByRef nYear as Long)
      Declare Function cmDaysFromJulian (ByVal nMonth as Short, _
                                         ByVal nDay as Short, _
                                         ByVal nYear as Long) as Long
      Declare Function cmJulianLeapYear(ByVal nYear as Long) as BOOLEAN
      
' Easter Support

      Declare Sub cmEasterCalculation(ByVal nRuleClass as Short, _
                                      ByVal nRule as Short, _
                                      ByVal nYear as Long, _
                                      ByRef nCalcDays as Long)
      Declare Function cmOrthodoxEasterDay (ByVal nYear as Long) as Long
      Declare Function cmChristianEasterDay (ByVal nYear as Long) as Long
      
' Chinese support

      Declare Function cmDaysFromChinese(ByVal nCycle as Short, _
                                         ByVal nYear as Long, _
                                         ByVal nMonth as Short, _
                                         ByVal bLeapMonth as BOOLEAN, _
                                         ByVal nDay as Short, _
                                         ByVal nCountry as Short) as Long
      Declare Function cmChineseDateCalculation(ByVal nRuleClass as Short, _
                                                ByVal nMonth as Short, _
                                                ByVal nDay as Short, _
                                                ByVal nGregorianYear as Long, _
                                                ByVal nRule as Short, _
                                                ByRef nCalcDays as Long) as BOOLEAN
      Declare Sub cmChineseFromDays(ByVal nDays as Long, _
                                    ByRef nCycle as Short, _
                                    ByRef nYear as Long, _
                                    ByRef nMonth as Short, _
                                    ByRef bLeapMonth as BOOLEAN, _
                                    ByRef nDay as Short, _
                                    ByVal nCountry as Short)
      Declare Sub cmChineseCycleAndYear(ByVal nGregorianYear as Long, _
                                        ByRef nCycle as Short, _
                                        ByRef nChineseYear as Long)
      Declare Function cmChineseNewYearOnOrBefore(ByVal nDays as Long, _
                                                  ByVal nCountry as Short) as Long
      Declare Function cmChineseNewYearInSui(ByVal nDays as Long, _
                                             ByVal nCountry as Short) as Long
      Declare Function cmChinesePriorLeapMonth(ByVal nMPrime as Long, _
                                               ByVal nM as Long, _
                                               ByVal nCountry as Short) as BOOLEAN
      Declare Function cmChineseNoMajorSolarTerm(ByVal nDays as Long, _
                                                 ByVal nCountry as Short) as BOOLEAN
      Declare Function cmChineseNewMoonBefore(ByVal nMoment as Double, _
                                              ByVal nCountry as Short) as Long
      Declare Function cmChineseNewMoonOnOrAfter(ByVal nMoment as Double, _
                                                 ByVal nCountry as Short) as Long
      Declare Function cmChineseWinterSolsticeOnOrBefore(ByVal nDays as Long, _
                                                         ByVal nCountry as Short) as Long
      Declare Function cmMinorSolarTermOnOrAfter(ByVal nMoment as Double, _
                                                 ByVal nCountry as Short) as Double
      Declare Function cmCurrentMinorSolarTerm(ByVal nMoment as Double, _
                                               ByVal nCountry as Short) as Long
      Declare Function cmMajorSolarTermOnOrAfter(ByVal nMoment as Double, _
                                                 ByVal nCountry as Short) as Double
      Declare Function cmMidnightInChina(ByVal nMoment as Double, _
                                         ByVal nCountry as Short) as Double
      Declare Function cmChineseSolarLongitudeOnOrAfter(ByVal nMoment as Double, _
                                                        ByVal nSolarTerm as Long, _
                                                        ByVal nCountry as Short) as Double
      Declare Function cmCurrentMajorSolarTerm(ByVal nMoment as Double, _
                                               ByVal nCountry as Short) as Long
      Declare Function cmChineseLocation(ByVal nMoment as Double, _
                                         ByVal nCountry as Short) as Double
      Declare Function cmChineseSexagesimalName(ByVal nYear as Long) as Short
      Declare Function cmChineseYearName(ByVal nYear as Long) as Short
      Declare Function cmChineseMonthName(ByVal nMonth as Short, _
                                          ByVal nYear as Long) as Short
      Declare Function cmChineseYearMarriageAuguries(ByVal nCycle as Short, _
                                                     ByVal nYear as Long, _
                                                     ByVal nCountry as Short) as Short
                                                     
' Hebrew support

      Declare Function cmHebrewDateCalculation(ByVal nMonth as Short, _
                                                ByVal nDay as Short, _
                                                ByVal nGregorianYear as Long, _
                                                ByVal nWeekday as Short, _
                                                ByVal nRule as Short, _
                                                ByRef nCalcDays as Long) as BOOLEAN
      Declare Function cmHebrewBirthDay(ByVal nBirthMonth as Short, _
                                        ByVal nBirthDay as Short, _
                                        ByVal nBirthYear as Long, _
                                        ByVal nForHebrewYear as Long) as Long
      Declare Sub cmHebrewFromDays(ByVal nDays as Long, _
                                   ByRef nMonth as Short, _
                                   ByRef nDay as Short, _
                                   ByRef nYear as Long)
      Declare Function cmDaysFromHebrew(ByVal nMonth as Short, _
                                        ByVal nDay as Short, _
                                        ByVal nYear as Long) as Long
      Declare Function cmLastDayOfHebrewMonth(ByVal nMonth as Long, _
                                              ByVal nYear as Long) as Long
      Declare Function cmLongMarheshvan(ByVal nYear as Long) as Long
      Declare Function cmShortKislev(ByVal nYear as Long) as BOOLEAN
      Declare Function cmDaysInHebrewYear(ByVal nYear as Long) as Long
      Declare Function cmHebrewNewYear(ByVal nYear as Long) as Long
      Declare Function cmHebrewYearLengthCorrection(ByVal nYear as Long) as Long
      Declare Function cmHebrewCalendarElapsedDays(ByVal nYear as Long) as Long
      Declare Function cmMolad(ByVal nMonth as Short, _
                               ByVal nYear as Long) as Double
      Declare Function cmLastMonthOfHebrewYear(ByVal nYear as Long) as Short
      Declare Function cmHebrewLeapYear(ByVal nYear as Long) as BOOLEAN
      Declare Function cmHebrewSabbaticalYear(ByVal nYear as Long) as BOOLEAN
      
' Islamic support

      Declare Function cmIslamicDateCalculation(ByVal nRule as Short, _
                                                ByVal nMonth as Short, _
                                                ByVal nDay as Short, _
                                                ByVal nGregorianYear as Long, _
                                                ByRef nHoliday1 as Long, _
                                                ByRef nHoliday2 as Long) as BOOLEAN
      Declare Sub cmIslamicInGregorian(ByVal nIslamicMonth as Short, _
                                       ByVal nIslamicDay as Short, _
                                       ByVal nGregorianYear as Long, _
                                       ByRef nFirstDate as Long, _
                                       ByRef bFirstValidDate as BOOLEAN, _
                                       ByRef nSecondDate as Long, _
                                       ByRef bSecondValidDate as BOOLEAN)
      Declare Sub cmIslamicFromDays(ByVal nDays as Long, _
                                    ByRef nMonth as Short, _
                                    ByRef nDay as Short, _
                                    ByRef nYear as Long)
      Declare Function cmDaysFromIslamic(ByVal nMonth as Short, _
                                         ByVal nDay as Short, _
                                         ByVal nYear as Long) as Long
      Declare Function cmPhasisOnOrBefore(ByVal nDays as Long) as Long
      
' Persian support

      Declare Function cmPersianDateCalculation(ByVal nMonth as Short, _
                                                ByVal nDay as Short, _
                                                ByVal nGregorianYear as Long, _
                                                ByRef nCalcDays as Long) as BOOLEAN
      Declare Function cmPersianYear(ByVal nGregorianYear as Long) as Long
      Declare Sub cmPersianFromDays(ByVal nDays as Long, _
                                    ByRef nMonth as Short, _
                                    ByRef nDay as Short, _
                                    ByRef nYear as Long)
      Declare Function cmDaysFromPersian(ByVal nMonth as Short, _
                                         ByVal nDay as Short, _
                                         ByVal nYear as Long) as Long
      Declare Function cmPersianNewYearOnOrBefore(ByVal nDays as Long) as Long

      Declare Function cmMiddayInTehran(ByVal nDays as Long) as Double
      
' Hindu Solar and Lunar support

      Declare Function cmHinduLunarDateCalculation(ByVal nRule as Short, _
                                                   ByVal nMonth as Short, _
                                                   ByVal nDay as Short, _
                                                   ByVal nGregorianYear as Long, _
                                                   ByRef nHoliday1 as Long, _
                                                   ByRef nHoliday2 as Long) as BOOLEAN
      Declare Function cmDaysFromHinduLunar(ByVal nMonth as Short, _
                                            ByVal bLeapMonth as BOOLEAN, _
                                            ByVal nDay as Short, _
                                            ByVal bLeapDay as BOOLEAN, _
                                            ByVal nYear as Long) as Long
      Declare Sub cmHinduLunarFromDays(ByVal nDays as Long, _
                                       ByRef nMonth as Short, _
                                       ByRef bLeapMonth as BOOLEAN, _
                                       ByRef nDay as Short, _
                                       ByRef bLeapDay as BOOLEAN, _
                                       ByRef nYear as Long)
      Declare Function cmHinduSolarDateCalculation(ByVal nRule as Short, _
                                                   ByVal nMonth as Short, _
                                                   ByVal nDay as Short, _
                                                   ByVal nGregorianYear as Long, _
                                                   ByRef nCalcDays as Long) as BOOLEAN
      Declare Function cmDaysFromHinduSolar(ByVal nMonth as Short, _
                                            ByVal nDay as Short, _
                                            ByVal nYear as Long) as Long
      Declare Sub cmHinduSolarFromDays(ByVal nDays as Long, _
                                       ByRef nMonth as Short, _
                                       ByRef nDay as Short, _
                                       ByRef nYear as Long)
      Declare Function cmExpunged(ByVal nMonth as Short, _
                                  ByVal nYear as Long) as BOOLEAN
      Declare Function cmAdjustedHindu(ByVal nMonth as Short, _
                                       ByVal bLeapMonth as BOOLEAN, _
                                       ByVal nDay as Short, _
                                       ByVal bLeapDay as BOOLEAN, _
                                       ByVal nYear as Long) as Long
      Declare Function cmAlmostEqual(ByVal nMonth1 as Short, _
                                     ByVal bLeapMonth1 as BOOLEAN, _
                                     ByVal nMonth2 as Short, _
                                     ByVal bLeapMonth2 as BOOLEAN) as Long
      Declare Function cmHinduLunarOnOrBefore(ByRef nMonth1 as Short, _
                                              ByRef bLeapMonth1 as BOOLEAN, _
                                              ByRef nDay1 as Short, _
                                              ByRef bLeapDay1 as BOOLEAN, _
                                              ByRef nYear1 as Long, _
                                              ByRef nMonth2 as Short, _
                                              ByRef bLeapMonth2 as BOOLEAN, _
                                              ByRef nDay2 as Short, _
                                              ByRef bLeapDay2 as BOOLEAN, _
                                              ByRef nYear2 as Long) as BOOLEAN
      Declare Function cmHinduSunRise(ByVal nDays as Long) as Double
      Declare Function cmHinduCalendarYear(ByVal nMoment as Double) as Long
      Declare Function cmHinduNewMoonBefore(ByVal nMoment as Double) as Double
      Declare Function cmHinduLunarDayFromMoment(ByVal nMoment as Double) as Long
      Declare Function cmHinduLunarPhase(ByVal nMoment as Double) as Double
      Declare Function cmHinduLunarLongitude(ByVal nMoment as Double) as Double
      Declare Function cmHinduZodiac(ByVal nMoment as Double) as Long
      Declare Function cmHinduSolarLongitudeAtOrAfter(ByVal nTargetLongitude as Double, _
                                                      ByVal nMoment as Double) as Double
      Declare Function cmHinduSolarLongitude(ByVal nMoment as Double) as Double
      Declare Function cmHinduTruePosition(ByVal nMoment as Double, _
                                           ByVal nPeriod as Double, _                
                                           ByVal nSize as Double, _                
                                           ByVal nAnomalistic as Double, _          
                                           ByVal nChange as Double) as Double
      Declare Function cmHinduMeanPosition(ByVal nMoment as Double, _
                                           ByVal nPeriod as Double) as Double
      Declare Function cmHinduArcsin(ByVal nAmp as Double) as Double
      Declare Function cmHinduSine(ByVal nDegrees as Double) as Double
      Declare Function cmHinduSineTable(ByVal nEntry as Long) as Double
      
' Coptic support

      Declare Function cmCopticDateCalculation(ByVal nMonth as Short, _
                                               ByVal nDay as Short, _
                                               ByVal nGregorianYear as Long, _
                                               ByRef nCalcDays as Long) as BOOLEAN
      Declare Sub cmCopticFromDays(ByVal nDays as Long, _
                                   ByRef nMonth as Short, _
                                   ByRef nDay as Short, _
                                   ByRef nYear as Long)
      Declare Function cmDaysFromCoptic(ByVal nMonth as Short, _
                                        ByVal nDay as Short, _
                                        ByVal nYear as Short) as Long
                                        
' Ethiopic support

      Declare Function cmEthiopicDateCalculation(ByVal nMonth as Short, _
                                                 ByVal nDay as Short, _
                                                 ByVal nGregorianYear as Long, _
                                                 ByRef nCalcDays as Long) as BOOLEAN
      Declare Sub cmEthiopicFromDays(ByVal nDays as Long, _
                                     ByRef nMonth as Short, _
                                     ByRef nDay as Short, _
                                     ByRef nYear as Long)
      Declare Function cmDaysFromEthiopic(ByVal nMonth as Short, _
                                          ByVal nDay as Short, _
                                          ByVal nYear as Short) as Long
                                          
' Roman support

      Declare Sub cmRomanFromDays(ByVal nDays as Long, _
                                  ByRef nMonth as Short, _
                                  ByRef nYear as Long, _
                                  ByRef nEvent as Short, _
                                  ByRef nCount as Short, _
                                  ByRef bLeap as BOOLEAN)
      Declare Function cmDaysFromRoman(ByVal nMonth as Short, _
                                       ByVal nYear as Long, _
                                       ByVal nEvent as Short, _
                                       ByVal nCount as Short, _
                                       ByVal bLeap as BOOLEAN) as Long
      Declare Function cmNonesOfMonth(ByVal nMonth as Short) as Short
      Declare Function cmIdesOfMonth(ByVal nMonth as Short) as Short
      
' Armenian support

      Declare Function cmDaysFromArmenian(ByVal nMonth as Short, _
                                          ByVal nDay as Short, _
                                          ByVal nYear as Long) as Long
      Declare Sub cmArmenianFromDays(ByVal nDays as Long, _
                                     ByRef nMonth as Short, _
                                     ByRef nDay as Short, _
                                     ByRef nYear as Long)
                                     
' Egyptian support
                                     
      Declare Function cmDaysFromEgyptian(ByVal nMonth as Short, _
                                          ByVal nDay as Short, _
                                          ByVal nYear as Long) as Long
      Declare Sub cmEgyptianFromDays(ByVal nDays as Long, _
                                     ByRef nMonth as Short, _
                                     ByRef nDay as Short, _
                                     ByRef nYear as Long)
                                     
' Bahai support

      Declare Function cmBahaiDateCalculation(ByVal nRule as Short, _
                                              ByVal nMonth as Short, _
                                              ByVal nDay as Short, _
                                              ByVal nGregorianYear as Long, _
                                              ByRef nCalcDays as Long) as BOOLEAN                                     
      Declare Sub cmBahaiFromDays(ByVal nDays as Long, _
                                  ByRef nMajor as Short, _
                                  ByRef nCycle as Short, _
                                  ByRef nMonth as Short, _
                                  ByRef nDay as Short, _
                                  ByRef nYear as Short)
      Declare Function cmDaysFromBahai(ByVal nMajor as Short, _
                                       ByVal nCycle as Short, _
                                       ByVal nYear as Short, _
                                       ByVal nMonth as Short, _
                                       ByVal nDay as Short) as Long
      Declare Function cmBahaiNewYearOnOrBefore(ByVal nDays as Long) as Double
      Declare Function cmBahaiSunset(ByVal nDays as Long) as Double
      
' Tibetan support

      Declare Function cmTibetanDateCalculation(ByVal nRule as Short, _
                                                ByVal nMonth as Short, _
                                                ByVal nDay as Short, _
                                                ByVal nGregorianYear as Long, _
                                                ByRef nCalcDays as Long) as BOOLEAN
      Declare Function cmTibetanNewYearInGregorian(ByVal nYear as Long) as Long
      Declare Function cmTibetanLosar(ByVal nYear as Long) as Long
      Declare Function cmTibetanLeapMonth(ByVal nMonth as Short, _
                                          ByVal nYear as Long) as BOOLEAN
      Declare Sub cmTibetanFromDays(ByVal nDays as Long, _
                                    ByRef nMonth as Short, _
                                    ByRef bLeapMonth as BOOLEAN, _
                                    ByRef nDay as Short, _
                                    ByRef bLeapDay as BOOLEAN, _
                                    ByRef nYear as Long)
      Declare Function cmDaysFromTibetan(ByVal nMonth as Short, _
                                         ByVal bLeapMonth as BOOLEAN, _
                                         ByVal nDay as Short, _
                                         ByVal bLeapDay as BOOLEAN, _
                                         ByVal nYear as Long) as Long
      Declare Function cmTibetanSunEquation(ByVal nAlpha as Double) as Double
      Declare Function cmTibetanMoonEquation(ByVal nAlpha as Double) as Double
      
' Astronomy support

      Declare Function cmLunarIllumination(byVal nMoment as Double) as Double
      Declare Sub cmMoonRiseAndSet(ByVal nSerial as LongInt, _
                                   ByVal bType as BOOLEAN, _
                                   ByRef uLocale as LOCATION_LOCALE, _
                                   arLunarTimes() as LUNAR_RISE_AND_SET)
      Declare Function cmEarthRadius(ByVal nLatitude as Double) as Double
      Declare Function cmSolarDistance(ByVal nMoment as Double) as Double
      Declare Function cmSolarEquationOfCenter(ByVal nC as Double) as Double
      Declare Function cmSunRise(ByVal nDays as Long, _
                                 ByVal nZone as Double, _
                                 ByVal nLatitude as Double, _
                                 ByVal nLongitude as Double, _
                                 ByVal nElevation as Double, _ 
                                 ByVal nDepression as Double, _                   
                                 ByRef bBogus as BOOLEAN) as Double
      Declare Function cmSunSet(ByVal nDays as Long, _
                                ByVal nZone as Double, _
                                ByVal nLatitude as Double, _
                                ByVal nLongitude as Double, _
                                ByVal nElevation as Double, _
                                ByVal nDepression as Double, _            
                                ByRef bBogus as BOOLEAN) as Double
      Declare Function cmDawn(ByVal nDays as Long, _
                              ByVal nZone as Double, _
                              ByVal nLatitude as Double, _
                              ByVal nLongitude as Double, _
                              ByVal nDepression as Double, _
                              ByRef bBogus as BOOLEAN) as Double
      Declare Function cmDusk(ByVal nDays as Long, _
                              ByVal nZone as Double, _
                              ByVal nLatitude as Double, _
                              ByVal nLongitude as Double, _
                              ByVal nDepression as Double, _
                              ByRef bBogus as BOOLEAN) as Double
      Declare Function cmMomentOfDepression(ByVal nApprox as Double, _
                                            ByVal nLatitude as Double, _
                                            ByVal nLongitude as Double, _
                                            ByVal nDepression as Double, _
                                            ByVal bEarly as BOOLEAN, _
                                            ByRef bBogus as BOOLEAN) as Double
      Declare Function cmApproxMomentOfDepression(ByVal nMoment as Double, _
                                                  ByVal nLatitude as Double, _
                                                  ByVal nLongitude as Double, _
                                                  ByVal nDepression as Double, _
                                                  ByVal bEarly as BOOLEAN, _
                                                  ByRef bBogus as BOOLEAN) as Double
      Declare Function cmSolarRefraction(ByVal nElevation as Double,ByVal nLatitude as Double) as Double
      Declare Function cmSineOffset(ByVal nMoment as Double, _
                                    ByVal nLatitude as Double, _
                                    ByVal nLongitude as Double, _
                                    ByVal nDepression as Double) as Double
      Declare Function cmLunarDistance(ByVal nMoment as Double) as Double
      Declare Function cmSumDistancePeriods(ByRef nE as Double, _
                                            ByRef nElongation as Double, _
                                            ByRef nSolarAnomaly as Double, _
                                            ByRef nLunarAnomaly as Double, _
                                            ByRef nMoonFromNode as Double, _
                                            ByVal nV as Double, _
                                            ByVal nW as Double, _
                                            ByVal nX as Double, _
                                            ByVal nY as Double, _
                                            ByVal nZ as Double) as Double
      Declare Function cmTopocentricLunarAltitude(ByVal nMoment as Double, _
                                                  ByRef uLocale as LOCATION_LOCALE) as Double
      Declare Function cmLunarParallax(ByVal nMoment as Double, _
                                       ByVal nLunarAltitude as Double, _
                                       ByVal nLatitude as Double, _
                                       ByVal nLongitude as Double) as Double
      Declare Function cmGeocentricLunarAltitude(ByVal nMoment as Double, _
                                                 ByRef uLocale as LOCATION_LOCALE) as Double
      Declare Function cmLunarLatitude(ByVal nMoment as Double) as Double
      Declare Function cmLunarPhaseAtOrBefore(ByVal nMoment as Double, _
                                              ByVal nTargetLongitude as Double) as Double
      Declare Function cmLunarPhaseAtOrAfter(ByVal nMoment as Double, _
                                             ByVal nTargetLongitude as Double) as Double
      Declare Function cmNewMoonAfter(ByVal nMoment as Double) as Double
      Declare Function cmNewMoonBefore(nMoment as Double) as Double
      Declare Function cmLunarPhase(ByVal nMoment as Double) as Double
      Declare Function cmLunarLongitude(nMoment as Double) as Double
      Declare Function cmSumLunarPeriods(ByRef nE as Double, _
                                         ByRef nElongation as Double, _
                                         ByRef nSolarAnomaly as Double, _
                                         ByRef nLunarAnomaly as Double, _
                                         ByRef nMoonFromNode as Double, _
                                         ByVal nV as Double, _
                                         ByVal nW as Double, _
                                         ByVal nX as Double, _
                                         ByVal nY as Double, _
                                         ByVal nZ as Double) as Double
      Declare Function cmMeanLunarLongitude(ByRef nC as Double) as Double
      Declare Function cmLunarElongation(ByRef nC as Double) as Double
      Declare Function cmSolarAnomaly(ByRef nC as Double) as Double
      Declare Function cmLunarAnomaly(ByRef nC as Double) as Double
      Declare Function cmMoonNode(ByRef nC as Double) as Double
      Declare Function cmNthNewMoon(nNthMoon as Long) as Double
      Declare Function cmCorrectionAdjustments(ByRef dtE as Double, _
                                               ByRef dtSolarAnomaly as Double, _
                                               ByRef dtLunarAnomaly as Double, _
                                               ByRef dtMoonArgument as Double, _
                                               ByVal dtV as Double, _
                                               ByVal dtW as Double, _
                                               ByVal dtX as Double, _
                                               ByVal dtY as Double, _
                                               ByVal dtZ as Double) as Double
      Declare Function cmAdditionalAdjustments(ByRef stK as Double, _
                                               ByVal stI as Double, _
                                               ByVal stJ as Double, _
                                               ByVal stL as Double) as Double
      Declare Function cmEstimatePriorSolarLongitude(ByVal nMoment as Double, _
                                                     ByVal nLongitude as Double) as Double
      Declare Function cmSeasonalEquinox(ByVal nYear as Long, _
                                         ByVal nEquinox as Long) as Double
      Declare Function cmSolarLongitudeAfter(ByVal nMoment as Double, _
                                             ByVal nTargetLongitude as Double) as Double
      Declare Function cmSolarLongitude(ByVal dtMoment as Double) as Double
      Declare Function cmSumSolarLongitudePeriods(ByRef dtC as Double, _
                                                  ByVal dwX as Double, _
                                                  ByVal dtY as Double, _
                                                  ByVal dtZ as Double) as Double
      Declare Function cmAberration(ByVal nC as Double) as Double
      Declare Function cmNutation (ByVal nC as Double) as Double
      Declare Function cmDeclination(ByVal nObliquity as Double, _
                                     ByVal nLatitude as Double, _
                                     ByVal nLongitude as Double) as Double
      Declare Function cmRightAscension(ByVal nObliquity as Double, _
                                        ByVal nLatitude as Double, _
                                        ByVal nLongitude as Double) as Double
      Declare Function cmSiderealFromMoment(ByVal nMoment as Double) as Double      
      Declare Function cmMeanTropicalYear(ByVal nC as Double) as Double
      Declare Function cmMidnight(ByVal nMoment as Double, _
                                  ByVal nZone as Double, _
                                  ByVal nLongitude as Double) as Double
      Declare Function cmMidday(ByVal nMoment as Double, _
                                ByVal nZone as Double, _
                                ByVal nLongitude as Double) as Double
      Declare Function cmApparentFromLocal(ByVal nMoment as Double, _
                                           ByVal nLongitude as Double) as Double
      Declare Function cmLocalFromApparent(ByVal nMoment as Double, _
                                           ByVal nLongitude as Double) as Double
      Declare Function cmEquationOfTime(ByVal nMoment as Double) as Double
      Declare Function cmSolarMeanAnomaly(ByVal nC as Double) as Double
      Declare Function cmEccentricityEarthOrbit(ByVal nCenturies as Double) as Double
      Declare Function cmObliquity (ByVal nCenturies as Double) as Double
      Declare Function cmAngle(ByVal nDegrees as Double, _
                               ByVal nMinutes as Double, _
                               ByVal nSeconds as Double) as Double
      Declare Function cmSinDegrees(ByVal nTheta as Double) as Double
      Declare Function cmCoSineDegrees(ByVal nTheta as Double) as Double
      Declare Function cmArcCoSineDegrees(ByVal nTheta as Double) as Double
      Declare Function cmArcSinDegrees(ByVal nTheta as Double) as Double
      Declare Function cmTangentDegrees(ByVal nTheta as Double) as Double
      Declare Function cmArcTanDegrees(ByVal nY as Double, _
                                       ByVal nX as Double) as Double
      Declare Function cmCalcDegrees(ByVal nDegrees as Double) as Double
      Declare Function cmJulianCenturies(ByVal nMoment as Double) as Double
      Declare Function cmDynamicalFromUniversal(ByVal nUniversal as Double) as Double
      Declare Function cmUniversalFromDynamical(ByVal nDynamical as Double) as Double
      Declare Function cmEphemerisCorrection (ByVal nMoment as Double) as Double
      Declare Function cmLocalFromStandard(ByVal nStandard as Double, _
                                           ByVal nZone as Double, _
                                           ByVal nLongitude as Double) as Double
      Declare Function cmStandardFromLocal(ByVal nLocal as Double, _
                                           ByVal nZone as Double, _
                                           ByVal nLongitude as Double) as Double
      Declare Function cmUniversalFromStandard(ByVal nStandard as Double, _
                                               ByVal nZone as Double) as Double
      Declare Function cmStandardFromUniversal(ByVal nUniversal as Double, _
                                               ByVal nZone as Double) as Double
      Declare Function cmLocalFromUniversal(ByVal nUniversal as Double, _
                                            ByVal nLongitude as Double) as Double
      Declare Function cmUniversalFromLocal(ByVal nLocal as Double, _
                                            ByVal nLongitude as Double) as Double
      Declare Function cmZoneFromLongitude(ByVal nLongitude as Double) as Double
      Declare Function cmDegreesToRadians(ByVal nDegrees as Double) as Double
      Declare Function cmRadiansToDegrees(ByVal nRadians as Double) as Double
      
' Common support      

      Declare Function cmDaylightSavings (ByVal nSerial as LongInt, _
                                          ByRef uLocation as LOCATION_LOCALE) as LongInt      
      Declare Sub cmTimeFromSerial(ByVal nTime as Long, _
                                   ByRef nHour as Short, _
                                   ByRef nMinute as Short, _
                                   ByRef nSecond as Short, _
                                   ByRef nMillisecond as Short)
      Declare Function cmTimeToSerial(ByVal nHour as Short, _
                                      ByVal nMinute as Short, _
                                      ByVal nSecond as Short, _
                                      ByVal nMillisecond as Short) as Long
      Declare Sub cmSerialBreakApart(ByRef nSerial as LongInt, _
                                     ByRef nDays as Long,      _
                                     ByRef nTime as Long)
      Declare Function cmAMod(ByVal x as Double, ByVal y as Double) as Double
      Declare Function cmMod(ByVal x as Double, ByVal y as Double) as Double
      Declare Function cmRound(ByVal x as Double) as Long
      Declare Function cmCeiling(ByVal x as Double) as Long
      Declare Function cmFloor(ByVal x as Double) as Long
      Declare Function cmSignum(ByVal nAny as Double) as Long
      Declare Sub cmSumOneMonth(ByVal nMonth as Long, _
                                uHistory as HISTORY_MONTHS, _
                                ByRef nSummary as Double)
      Declare Sub cmClearSummary(arSummary() as Double)
      Declare Sub cmShiftHistory(ByVal nMonth as Short, _
                                 uHistory as HISTORY_MONTHS)
      Declare Function cmMomentToSerial(ByRef nMoment as Double) as LongInt
      Declare Function cmSerialToMoment(ByVal nSerial as LongInt) as Double

   Public:

      Declare Constructor
      Declare Destructor

' Date Calculation Interface

      Declare Sub EraseNonBusinessDates()
      Declare Sub DateCalculation(arRules() as DATE_CALCULATION)
      Declare Sub SetBusinessWeekday(ByVal nWeekday as Short, _
                                     ByVal bWorkday as BOOLEAN)
      Declare Sub GetSavedNonBusinessDates(arDates() as LongInt)
      Declare Function IsBusinessDay(ByVal nSerial as LongInt) as BOOLEAN
      Declare Property GetBusinessWeekday(ByVal nWeekday as Short) as BOOLEAN
      Declare Property NonBusinessDatesLimit() as Long
      Declare Property NonBusinessDatesSaved() as Long
      Declare Function BusinessDaysBetween(ByVal nStart as LongInt, _
                                           ByVal nEnd as LongInt) as Long
      Declare Function NonBusinessDaysBetween(ByVal nStart as LongInt, _
                                              ByVal nEnd as LongInt) as Long
      Declare Function BusinessMonthBegin (ByVal nMonth as Long, _
                                           ByVal nYear as Long) as LongInt
      Declare Function BusinessMonthEnd(ByVal nMonth as Long, _
                                        ByVal nYear as Long) as LongInt
      Declare Function AddBusinessDays(ByVal nSerial as LongInt, _
                                       ByVal nDays as Long) as LongInt
      
' Gregorian Interface

      Declare Function ValidGregorianDate(ByRef uDate as GREGORIAN_DATE) as Short      
      Declare Sub GregorianFromSerial(ByVal nSerial as LongInt, _
                                      ByRef uDate as GREGORIAN_DATE)
      Declare Function SerialFromGregorian(ByRef uDate as GREGORIAN_DATE) as LongInt
      Declare Function GregorianDaysRemaining(ByVal nMonth as Short, _
                                              ByVal nDay as Short, _
                                              ByVal nYear as Long) as Short
      Declare Function GregorianDayNumber(ByVal nMonth as Short, _
                                          ByVal nDay as Short, _
                                          ByVal nYear as Long) as Short
      Declare Function GregorianDaysInMonth(ByVal nMonth as Short, _
                                            ByVal nYear as Long) as Short
      Declare Function GregorianFiscalQuarter(ByVal nCalendarMonth as Short, _
                                              ByVal nFiscalYearBeginsMonth as Short) as Short
      Declare Function GregorianFiscalMonth(ByVal nCalendarMonth as Short, _
                                            ByVal nFiscalYearBeginsMonth as Short) as Short
      Declare Sub GregorianFiscalMonthYears(ByVal nCalendarYear as Long, _
                                            ByVal nFiscalYearBeginsMonth as Short, _
                                            ByVal bUseEndingMonth as BOOLEAN,_
                                            arFiscalYears() as Long)
      Declare Function GregorianFiscalYear(ByVal nCalendarMonth as Short, _
                                           ByVal nCalendarYear as Long, _
                                           ByVal nFiscalYearBeginsMonth as Short, _
                                           ByVal bUseEndingMonth as BOOLEAN) as Long
                                           
' Rolling 13 Month History Interface

      Declare Function UpdateHistory(ByRef nHistoryMonth as Short, _
                                     ByRef nHistoryYear as Long, _
                                     arHistory() as HISTORY_MONTHS, _
                                     ByVal nUpdateMonth as Short, _
                                     ByVal nUpdateYear as Long, _
                                     arUpdates() as Double) as BOOLEAN
      Declare Sub HistoryMonthsAndYears(ByVal nHistoryMonth as Long, _
                                        ByVal nHistoryYear as Long, _
                                        ByVal nFiscalYearBeginsMonth as Short, _
                                        ByVal bUseEndingMonth as BOOLEAN, _
                                        arFiscalYears() as Long)
      Declare Function SummarySameMonthLastYear(ByVal nHistoryMonth as Short, _
                                                arHistory() as HISTORY_MONTHS, _
                                                arSummary() as Double) as BOOLEAN
      Declare Function SummaryLastQuarter(ByVal nHistoryMonth as Short, _
                                          ByVal nStartFiscalMonth as Short, _
                                          arHistory() as HISTORY_MONTHS, _
                                          arSummary() as Double) as BOOLEAN
      Declare Function SummaryQuarterToDate(ByVal nHistoryMonth as Short, _
                                            ByVal nStartFiscalMonth as Long, _
                                            arHistory() as HISTORY_MONTHS, _
                                            arSummary() as Double) as BOOLEAN
      Declare Function SummaryYearToDate(ByVal nHistoryMonth as Short, _
                                         ByVal nStartFiscalMonth as Short, _
                                         arHistory() as HISTORY_MONTHS, _
                                         arSummary() as Double) as Long
                                         
' Time Interface

      Declare Function SerialAddTime(ByVal nSerial as LongInt, _
                                     ByRef uTime as TIME_DURATION) as LongInt
      Declare Function SerialSubtractTime(ByVal nSerial as LongInt, _
                                          ByRef uTime as TIME_DURATION) as LongInt
      Declare Sub SerialDifference(ByVal nFromSerial as LongInt, _
                                   ByVal nToSerial as LongInt, _
                                   ByRef uDiff as TIME_DURATION)
      Declare Function TimeToSerial(ByRef uTime as TIME_DURATION) as LongInt
      Declare Sub SerialToTime(ByVal nSerial as LongInt, _
                               ByRef uTime as TIME_DURATION)

' Unix Interface

      Declare Function UnixFromSerial(ByVal nSerial as LongInt) as LongInt
      Declare Function SerialFromUnix(ByVal nUnix as LongInt) as LongInt

' Excel Interface

      Declare Function ExcelFromSerial(ByVal nSerial as LongInt, _
                                       ByVal bBaseYear1904 as BOOLEAN) as Double
      Declare Function SerialFromExcel(ByVal nExcel as Double, _
                                       ByVal bBaseYear1904 as BOOLEAN) as LongInt
                                       
' Julian Interface

      Declare Function ValidJulianDate(ByRef uDate as JULIAN_DATE) as Short
      Declare Sub JulianFromSerial(ByVal nSerial as LongInt, _
                                   ByRef uDate as JULIAN_DATE)
      Declare Function SerialFromJulian(ByRef uDate as JULIAN_DATE) as LongInt
      
' ISO Interface

      Declare Function ValidISODate(ByRef uDate as ISO_DATE) as Short
      Declare Sub ISOFromSerial(ByVal nSerial as LongInt, _
                                ByRef uDate as ISO_DATE)
      Declare Function SerialFromISO(ByRef uDate as ISO_DATE) as LongInt
      
' Chinese Interface

      Declare Function ValidChineseDate(ByRef uDate as CHINESE_DATE) as Short
      Declare Sub ChineseFromSerial(ByVal nSerial as LongInt, _
                                    ByRef uDate as CHINESE_DATE)
      Declare Function SerialFromChinese(ByRef uDate as CHINESE_DATE) as LongInt
      
' Hebrew Interface

      Declare Function ValidHebrewDate(ByRef uDate as HEBREW_DATE) as Short
      Declare Sub HebrewFromSerial(ByVal nSerial as LongInt, _
                                   ByRef uDate as HEBREW_DATE)
      Declare Function SerialFromHebrew(ByRef uDate as HEBREW_DATE) as LongInt
      Declare Sub HebrewBirthDay(ByRef uBirth as HEBREW_DATE, _
                                 ByVal nForHebrewYear as Long, _
                                 ByRef uBirthDay as HEBREW_DATE)
      Declare Sub HebrewBirthDayInGregorian(ByRef uBirth as HEBREW_DATE, _
                                            ByVal nForGregorianYear as Long, _
                                            ByRef uBirthDay1 as GREGORIAN_DATE, _
                                            ByRef uBirthDay2 as GREGORIAN_DATE)
                                            
' Islamic Interface

      Declare Function ValidIslamicDate(ByRef uDate as ISLAMIC_DATE) as Short
      Declare Sub IslamicFromSerial(ByVal nSerial as LongInt, _
                                    ByRef uDate as ISLAMIC_DATE)
      Declare Function SerialFromIslamic(ByRef uDate as ISLAMIC_DATE) as LongInt
      
' Persian Interface

      Declare Function ValidPersianDate(ByRef uDate as PERSIAN_DATE) as Short
      Declare Sub PersianFromSerial(ByVal nSerial as LongInt, _
                                    ByRef uDate as PERSIAN_DATE)
      Declare Function SerialFromPersian(ByRef uDate as PERSIAN_DATE) as LongInt
      
' Hindu Interface

      Declare Function ValidHinduSolarDate(ByRef uDate as HINDU_SOLAR_DATE) as Short
      Declare Function ValidHinduLunarDate(ByRef uDate as HINDU_LUNAR_DATE) as Short
      Declare Sub HinduLunarFromSerial(ByVal nSerial as LongInt, _
                                       ByRef uDate as HINDU_LUNAR_DATE)
      Declare Function SerialFromHinduLunar(ByRef uDate as HINDU_LUNAR_DATE) as LongInt
      Declare Sub HinduSolarFromSerial(ByVal nSerial as LongInt, _
                                       ByRef uDate as HINDU_SOLAR_DATE)
      Declare Function SerialFromHinduSolar(ByRef uDate as HINDU_SOLAR_DATE) as LongInt
                                            
' Coptic Interface

      Declare Function ValidCopticDate(ByRef uDate as COPTIC_DATE) as Short
      Declare Sub CopticFromSerial(ByVal nSerial as LongInt, _
                                   ByRef uDate as COPTIC_DATE)
      Declare Function SerialFromCoptic(ByRef uDate as COPTIC_DATE) as LongInt
      
' Ethiopic Interface

      Declare Function ValidEthiopicDate(ByRef uDate as ETHIOPIC_DATE) as Short
      Declare Sub EthiopicFromSerial(ByVal nSerial as LongInt, _
                                     ByRef uDate as ETHIOPIC_DATE)
      Declare Function SerialFromEthiopic(ByRef uDate as ETHIOPIC_DATE) as LongInt

' Roman Interface

      Declare Function ValidRomanDate(ByRef uDate as ROMAN_DATE) as Short      
      Declare Sub RomanFromSerial(ByVal nSerial as LongInt, _
                                  ByRef uDate as ROMAN_DATE)
      Declare Function SerialFromRoman(ByRef uDate as ROMAN_DATE) as LongInt
      
' Egyptian Interface

      Declare Function ValidEgyptianDate(ByRef uDate as EGYPTIAN_DATE) as Short
      Declare Sub EgyptianFromSerial(ByVal nSerial as LongInt, _
                                     ByRef uDate as EGYPTIAN_DATE)
      Declare Function SerialFromEgyptian(ByRef uDate as EGYPTIAN_DATE) as LongInt
      
' Armenian Interface

      Declare Function ValidArmenianDate(ByRef uDate as ARMENIAN_DATE) as Short      
      Declare Sub ArmenianFromSerial(ByVal nSerial as LongInt, _
                                     ByRef uDate as ARMENIAN_DATE)
      Declare Function SerialFromArmenian(ByRef uDate as ARMENIAN_DATE) as LongInt
      
' Bahai Interface

      Declare Function ValidBahaiDate(ByRef uDate as BAHAI_DATE) as Short      
      Declare Sub BahaiFromSerial(ByVal nSerial as LongInt, _
                                  ByRef uDate as BAHAI_DATE)
      Declare Function SerialFromBahai(ByRef uDate as BAHAI_DATE) as LongInt
      
' Tibetan Interface

      Declare Function ValidTibetanDate(ByRef uDate as TIBETAN_DATE) as Short
      Declare Sub TibetanFromSerial(ByVal nSerial as LongInt, _
                                    ByRef uDate as TIBETAN_DATE)
      Declare Function SerialFromTibetan(ByRef uDate as TIBETAN_DATE) as LongInt      
      
' Astronomy Interface

      Declare Sub LunarIllumination(ByVal nSerial as LongInt, _
                                    ByVal nZoneHours as Double, _
                                    ByRef nIllumination as Double, _
                                    ByRef bWaxing as BOOLEAN, _
                                    ByRef bCrescent as BOOLEAN)                                         
      Declare Sub LunarRiseAndSet(ByVal nFromSerial as LongInt, _
                                  ByVal nToSerial as LongInt, _
                                  ByVal bType as BOOLEAN, _
                                  ByRef uLocale as LOCATION_LOCALE, _
                                  arLunarTimes() as LUNAR_RISE_AND_SET)
      Declare Function SunTransit(ByRef uLocale as LOCATION_LOCALE, _
                                  ByVal nSerial as LongInt) as LongInt
      Declare Sub Sunrise(ByRef uLocale as LOCATION_LOCALE, _
                          ByRef nSunrise as LongInt, _
                          ByVal nDepression as Double, _
                          ByRef bBogus as BOOLEAN, _
                          ByVal nSerial as LongInt)
      Declare Sub Sunset(ByRef uLocale as LOCATION_LOCALE, _
                         ByRef nSunset as LongInt, _
                         ByVal nDepression as Double, _
                         ByRef bBogus as BOOLEAN, _
                         ByVal nSerial as LongInt)
      Declare Function LunarDistance(ByVal nSerial as LongInt) as Double
      Declare Function SolarDistance(ByVal nSerial as LongInt) as Double
      Declare Sub LunarPhases(ByVal nFromSerial as LongInt, _
                              ByVal nToSerial as LongInt, _
                              ByRef uLocale as LOCATION_LOCALE, _
                              arPhases() as LUNAR_PHASES)
      Declare Function SeasonalEquinox(ByRef uLocale as LOCATION_LOCALE, _
                                       ByVal nYear as Long, _
                                       ByVal nEquinox as Long) as LongInt
      Declare Function DecimalDegrees(ByVal nDegrees as Double, _
                                      ByVal nMinutes as Double, _
                                      ByVal nSeconds as Double) as Double
      Declare Function GreatCircleDistance(ByRef uFrom as LOCATION_LOCALE, _
                                           ByRef uTo as LOCATION_LOCALE) as Double
      Declare Function CompassBearing(ByRef uFrom as LOCATION_LOCALE, _
                                      ByRef uTo as LOCATION_LOCALE) as Double
      Declare Function KilometersToMiles(ByVal nKilometers as Double) as Double
      Declare Function MilesToKilometers(ByVal nMiles as Double) as Double
      Declare Function KilometersToNauticalMiles(ByVal nKilometers as Double) as Double
      Declare Function NauticalMilesToKilometers(ByVal nNauticalMiles as Double) as Double
      
' Development Interface

      Declare Sub Develop()
   
End Type

Constructor cCalendar()

    This.iUBound = -1
    
' For business date calculations, set default Mon-Fri workweek    
    
    This.arBusinessWeekdays(cCalendarClass.MONDAY) = True
    This.arBusinessWeekdays(cCalendarClass.TUESDAY) = True
    This.arBusinessWeekdays(cCalendarClass.WEDNESDAY) = True
    This.arBusinessWeekdays(cCalendarClass.THURSDAY) = True
    This.arBusinessWeekdays(cCalendarClass.FRIDAY) = True
    This.arBusinessWeekdays(cCalendarClass.SATURDAY) = False
    This.arBusinessWeekdays(cCalendarClass.SUNDAY) = False

End Constructor
Destructor cCalendar()

End Destructor

' ########################################################################################
' Development
' ########################################################################################

' ========================================================================================
' Stub for development and testing of new features
' ========================================================================================
Private Sub cCalendar.Develop()

End Sub
 
' ########################################################################################
' Astronomy
' ########################################################################################

' ========================================================================================
' Convert Kilometers to Miles
' ========================================================================================
Private Function cCalendar.KilometersToMiles (ByVal nKilometers as Double) as Double

     Function = nKilometers * .621371

End Function
' ========================================================================================
' Convert Miles To Kilometers
' ========================================================================================
Private Function cCalendar.MilesToKilometers (ByVal nMiles as Double) as Double

     Function = nMiles / .621371

End Function
' ========================================================================================
' Convert Kilometers To Nautical Miles
' ========================================================================================
Private Function cCalendar.KilometersToNauticalMiles (ByVal nKilometers as Double) as Double

     Function = nKilometers * .539957

End Function
' ========================================================================================
' Convert Nautical Miles To Kilometers
' ========================================================================================
Private Function cCalendar.NauticalMilesToKilometers (ByVal nNauticalMiles as Double) as Double

     Function = nNauticalMiles / .539957

End Function
' ========================================================================================
' Calculate distance in Great Circle Kilometers between two points
' ========================================================================================
Private Function cCalendar.GreatCircleDistance (ByRef uFrom as LOCATION_LOCALE, _
                                                ByRef uTo as LOCATION_LOCALE) as Double

Dim nLongitudeDifference   as Double
Dim nLatitudeDifference    as Double
Dim nA                     as Double
Dim nC                     as Double

    nLongitudeDifference = cmDegreesToRadians(uTo.Longitude - uFrom.Longitude)
    nLatitudeDifference = cmDegreesToRadians(uTo.Latitude - uFrom.Latitude)

    nA = Sqr(Sin(nLatitudeDifference / 2)^2 _
        + Cos(cmDegreesToRadians(uFrom.Latitude)) _
        * Cos(cmDegreesToRadians(uTo.Latitude)) _
        * Sin(nLongitudeDifference / 2)^2)

    nC = 2.0 * Asin(IIf(1 > nA,nA,1))

    Function = 6371.01 * nC

End Function
' ========================================================================================
' Give a source location, return the compass direction of the destination location
' ========================================================================================
Private Function cCalendar.CompassBearing (ByRef uFrom as LOCATION_LOCALE, _
                                           ByRef uTo as LOCATION_LOCALE) as Double

Dim nX             as Double
Dim nY             as Double 

    nY = cmSinDegrees(uTo.Longitude - uFrom.Longitude)
    nX = (cmCoSineDegrees(uFrom.Latitude) * cmTangentDegrees(uTo.Latitude)) - cmSinDegrees(uFrom.Latitude) * cmCoSineDegrees(uFrom.Longitude - uTo.Longitude)
    
    If (nX = 0 AndAlso nY = 0) OrElse uTo.Latitude = 90 Then
    
       Function = 0
       
    Else
    
       If uTo.Latitude = -90 Then
       
          Function = 180
          
       Else
       
          Function = cmArcTanDegrees(nY,nX)
          
       End If
       
    End If 

End Function
' ========================================================================================
' Calculate Transit of the Sun at location Locale. 
' ========================================================================================
Private Function cCalendar.SunTransit (ByRef uLocale as LOCATION_LOCALE, _
                                       ByVal nSerial as LongInt) as LongInt

    Function = cmMomentToSerial(cmMidday(cmFloor(nSerial / cCalendarClass.ONE_DAY), _
                                         uLocale.Zone,uLocale.Longitude))

End Function
' ========================================================================================
' Calculate Sunrise at location Locale.
' ========================================================================================
Private Sub cCalendar.Sunrise (ByRef uLocale as LOCATION_LOCALE, _
                               ByRef nSunrise as LongInt, _
                               ByVal nDepression as Double, _
                               ByRef bBogus as BOOLEAN, _
                               ByVal nSerial as LongInt)

' If Sunrise does not occur (usually at far northern/southern latitudes) return is bBogus is TRUE, FALSE otherwise

Dim nMoment      as Double

    nMoment = cmFloor(nSerial / cCalendarClass.ONE_DAY)

    nMoment = cmSunRise(nMoment, _
                        uLocale.Zone, _
                        uLocale.Latitude, _
                        uLocale.Longitude, _
                        uLocale.Elevation, _
                        nDepression, _
                        bBogus)

' If event was found, return time

    Select Case bBogus

    Case False

        nSunrise = cmMomentToSerial(nMoment)
        
' Round to nearest minute

        nSunrise = cmDaylightSavings(((cmFloor((nSunrise + (30 * cCalendarClass.ONE_SECOND)) _
                 / cCalendarClass.ONE_MINUTE)) * cCalendarClass.ONE_MINUTE),uLocale)

    End Select

End Sub
' ========================================================================================
' Calculate Sunset at location Locale. 
' ========================================================================================
Private Sub cCalendar.Sunset (ByRef uLocale as LOCATION_LOCALE, _
                              ByRef nSunset as LongInt, _
                              ByVal nDepression as Double, _
                              ByRef bBogus as BOOLEAN, _
                              ByVal nSerial as LongInt)

' If Sunset does not occur (usually at far northern/southern latitudes) return is bBogus is TRUE, FALSE otherwise

Dim nMoment      as Double

    nMoment = cmFloor(nSerial / cCalendarClass.ONE_DAY)

    nMoment = cmSunSet(nMoment, _
                       uLocale.Zone, _
                       uLocale.Latitude, _
                       uLocale.Longitude, _
                       uLocale.Elevation, _
                       nDepression, _
                       bBogus)

' If event was found, return time

    Select Case bBogus

    Case False

        nSunset = cmMomentToSerial(nMoment)
        
' Round to nearest minute

        nSunset = cmDaylightSavings(((cmFloor((nSunset + (30 * cCalendarClass.ONE_SECOND)) _
                / cCalendarClass.ONE_MINUTE)) * cCalendarClass.ONE_MINUTE),uLocale)

    End Select

End Sub
' ========================================================================================
' Lunar Illumination
' ========================================================================================
Private Sub cCalendar.LunarIllumination (ByVal nSerial as LongInt, _
                                         ByVal nZoneHours as Double, _     ' < 0 for eastern longitudes
                                         ByRef nIllumination as Double, _
                                         ByRef bWaxing as BOOLEAN, _
                                         ByRef bCrescent as BOOLEAN)

' Returns the percent illumination for the moon at a given date/time and timezone
' True/False flags are also returned crescent/waning/waxing
' bWaxing = %TRUE or %FALSE (for waning moons)

Dim nMoment    as Double
Dim nToday     as Double
Dim nTomorrow  as Double


    nMoment = cmDynamicalFromUniversal(cmUniversalFromStandard(cmSerialToMoment(nSerial),nZoneHours))
    nToday = cmLunarIllumination(nMoment)

' Use next day illumination to determine waning/waxing

    nTomorrow = cmLunarIllumination(nMoment + 1)

    bCrescent = IIF(nToday <= .5,True,False)
    bWaxing = IIF(nTomorrow > nToday,True,False)

    nIllumination = nToday * 100

End Sub 
' ========================================================================================
' Lunar Distance in Kilometers
' ========================================================================================
Private Function cCalendar.LunarDistance (ByVal nSerial as LongInt) as Double
   
    Function = cmLunarDistance(cmSerialToMoment(nSerial)) / 1000   

End Function
' ========================================================================================
' Solar Distance in Kilometers
' ========================================================================================
Private Function cCalendar.SolarDistance (ByVal nSerial as LongInt) as Double
   
    Function = cmSolarDistance(cmSerialToMoment(nSerial)) * 149597870.691   

End Function
' ========================================================================================
' Moonrise and Moonset Local Times
' ========================================================================================
Private Sub cCalendar.LunarRiseAndSet (ByVal nFromSerial as LongInt, _
                                       ByVal nToSerial as LongInt, _
                                       ByVal bType as BOOLEAN, _
                                       ByRef uLocale as LOCATION_LOCALE, _
                                       arLunarTimes() as LUNAR_RISE_AND_SET)
Dim nLunarDay     as LongInt

' Remove time from requested dates

    nFromSerial = cmMomentToSerial(cmFloor(cmSerialToMoment(nFromSerial)))
    nToSerial = cmMomentToSerial(cmFloor(cmSerialToMoment(nToSerial)))
    
    If nFromSerial > nToSerial Then
    
       Swap nFromSerial, nToSerial
       
    End If
    
    Erase arLunarTimes
    
' Loop through date range

    For nLunarDay = nFromSerial To nToSerial Step cCalendarClass.ONE_DAY
    
       cmMoonRiseAndSet(nLunarDay,bType,uLocale,arLunarTimes())    
    
    Next         
 
End Sub
' ========================================================================================
' Lunar Phases
' ========================================================================================
Private Sub cCalendar.LunarPhases (ByVal nFromSerial as LongInt, _
                                   ByVal nToSerial as LongInt, _
                                   ByRef uLocale as LOCATION_LOCALE, _
                                   arPhases() as LUNAR_PHASES)

' Given a starting and ending date, return the major moon phases for the period

Dim nFrom             as Long
Dim nTo               as Long
Dim bLoop             as BOOLEAN
Dim nMoon             as Double
Dim iIndex            as Long
Dim arWork()          as LUNAR_PHASES 

    If nFromSerial > nToSerial Then
    
       Swap nFromSerial, nToSerial
      
    End If
   
    Erase arPhases
    
    nFrom = cmFloor(nFromSerial / cCalendarClass.ONE_DAY)
    nTo = cmFloor(nToSerial / cCalendarClass.ONE_DAY)
    
' Cycle through the period from end to beginning 28 days at a time
' and check for all 4 major lunar phases in each period

    bLoop = True
    
    While bLoop = True
       
' Check for New Moon
    
        nMoon = cmLunarPhaseAtOrBefore(nTo + 1,cCalendarClass.NEWMOON)
    
        If nMoon >= nFrom Then

           nMoon = cmStandardFromUniversal(nMoon,uLocale.Zone)
           
           If cmFloor(nMoon) >= nFrom AndAlso cmFloor(nMoon) <= nTo Then
           
              ReDim Preserve arWork(0 To UBound(arWork) + 1)
              arWork(UBound(arWork)).LunarSerialTime = cmMomentToSerial(nMoon)
              arWork(UBound(arWork)).Phase = cCalendarClass.NEWMOON
              
           End If    
       
        End If
        
' Check for First Quarter Moon
    
        nMoon = cmLunarPhaseAtOrBefore(nTo + 1,cCalendarClass.FIRSTQUARTERMOON)
    
        If nMoon >= nFrom Then

           nMoon = cmStandardFromUniversal(nMoon,uLocale.Zone)
           
           If cmFloor(nMoon) >= nFrom AndAlso cmFloor(nMoon) <= nTo Then           

              ReDim Preserve arWork(0 To UBound(arWork) + 1)           
              arWork(UBound(arWork)).LunarSerialTime = cmMomentToSerial(nMoon)
              arWork(UBound(arWork)).Phase = cCalendarClass.FIRSTQUARTERMOON
              
           End If    
    
        End If
        
' Check for Last Quarter Moon
    
        nMoon = cmLunarPhaseAtOrBefore(nTo + 1,cCalendarClass.LASTQUARTERMOON)
    
        If nMoon >= nFrom Then

           nMoon = cmStandardFromUniversal(nMoon,uLocale.Zone)
           
           If cmFloor(nMoon) >= nFrom AndAlso cmFloor(nMoon) <= nTo Then
                      
              ReDim Preserve arWork(0 To UBound(arWork) + 1)
              arWork(UBound(arWork)).LunarSerialTime = cmMomentToSerial(nMoon)
              arWork(UBound(arWork)).Phase = cCalendarClass.LASTQUARTERMOON
              
           End If    
    
        End If
        
' Check for Full Moon
    
        nMoon = cmLunarPhaseAtOrBefore(nTo + 1,cCalendarClass.FULLMOON)
    
        If nMoon >= nFrom Then

           nMoon = cmStandardFromUniversal(nMoon,uLocale.Zone)
           
           If cmFloor(nMoon) >= nFrom AndAlso cmFloor(nMoon) <= nTo Then
                      
              ReDim Preserve arWork(0 To UBound(arWork) + 1)
              arWork(UBound(arWork)).LunarSerialTime = cmMomentToSerial(nMoon)
              arWork(UBound(arWork)).Phase = cCalendarClass.FULLMOON
              
           End If    
    
        End If
    
        nTo = nTo - 28
        bLoop = IIf(nFrom <= nTo,True,False)    
    
    Wend
    
' Sort the list. A large list is not expected so a simple swap sort should be ok

    If UBound(arWork) >= 0 Then
    
       bLoop = True
    
       While bLoop = True

        bloop = False
        
        For iIndex = 1 To UBound(arWork)
        
            If cmFloor(cmSerialToMoment(arWork(iIndex - 1).LunarSerialTime)) > cmFloor(cmSerialToMoment(arWork(iIndex).LunarSerialTime)) Then
            
               Swap arWork(iIndex - 1), arWork(iIndex)       
               bloop = True
               
            End If
            
        Next
    
       Wend
      
    End If
    
' Save phases found and remove duplicates

    If UBound(arWork) >= 0 Then
    
' Save first phase found
    
       ReDim arPhases(0 To 0)
       arPhases(0) = arWork(0)
       
' Adjust for Daylight Savings

       arPhases(0).LunarSerialTime = cmDayLightSavings(arPhases(0).LunarSerialTime,uLocale)                
       
' Check and save the remainder of the phases 
    
       If UBound(arWork) >= 0 Then
       
          For iIndex = 1 To UBound(arWork) 
    
             If cmFloor(cmSerialToMoment(arWork(iIndex - 1).LunarSerialTime)) <> cmFloor(cmSerialToMoment(arWork(iIndex).LunarSerialTime)) Then
           
' Adjust for Daylight Savings

                arWork(iIndex).LunarSerialTime = cmDayLightSavings(arWork(iIndex).LunarSerialTime,uLocale)                
           
                ReDim Preserve arPhases(0 To UBound(arPhases) + 1)
                arPhases(UBound(arPhases)) = arWork(iIndex)
                
            End If
            
         Next     
      
       End If
       
    End If
                
End Sub
' ========================================================================================
' Seasonal Equinox
' ========================================================================================
Private Function cCalendar.SeasonalEquinox (ByRef uLocale as LOCATION_LOCALE, _
                                            ByVal nYear as Long, _
                                            ByVal nEquinox as Long) as LongInt

' Get one of the seasonal equinox (Spring,Summer,Autumn,Winter)
' nYear is a Gregorian Year. Serial Date and Time is return value in uLocale.Zone.

Dim nMoment      as Double

' Get equinox as a UTC Moment

    nMoment = cmSeasonalEquinox(nYear,nEquinox)
    
' Return in local time

    nMoment = cmStandardFromUniversal(nMoment,uLocale.Zone)
    
    Function = cmDaylightSavings(cmMomentToSerial(nMoment),uLocale)

End Function
' ========================================================================================
' Decimal Degrees
' ========================================================================================
Private Function cCalendar.DecimalDegrees (ByVal nDegrees as Double, _
                                           ByVal nMinutes as Double, _
                                           ByVal nSeconds as Double) as Double
                        
' Convert separate Degrees, Minutes, and, Seconds to decimal form

    Function = cmAngle(nDegrees,nMinutes,nSeconds)

End Function

' ########################################################################################
' Business Day Calculations
' ########################################################################################

' ========================================================================================
' Add Business Days to Date
' ========================================================================================
Private Function cCalendar.AddBusinessDays (ByVal nSerial as LongInt, _
                                            ByVal nDays as Long) as LongInt
                                   
' nDays < 0 subtract days

Dim nSign               as Long
Dim nBusinessDays       as Long
Dim bLoop               as BOOLEAN

    If nDays = 0 Then
    
       Function = nSerial
       Exit Function
       
    End If
    
' We need one day of the week to be a business day to avoid hanging in our loop

    If This.arBusinessWeekdays(cCalendarClass.MONDAY) = False AndAlso This.arBusinessWeekdays(cCalendarClass.TUESDAY) = False _
       AndAlso This.arBusinessWeekdays(cCalendarClass.WEDNESDAY) = False AndAlso This.arBusinessWeekdays(cCalendarClass.THURSDAY) = False _
       AndAlso This.arBusinessWeekdays(cCalendarClass.FRIDAY) = False AndAlso This.arBusinessWeekdays(cCalendarClass.SATURDAY) = False _
       AndAlso This.arBusinessWeekdays(cCalendarClass.SUNDAY) = False Then
 
       Function = nSerial
       Exit Function         
          
    End If

    nSign = cmSignum(nDays)
    nDays = Abs(nDays)
    bLoop = True
    nBusinessDays = 0

    While bLoop = True
    
          nSerial = nSerial + cCalendarClass.ONE_DAY * nSign
          
          If IsBusinessDay(nSerial) = True Then
          
             nBusinessDays = nBusinessDays + 1
             
          End If
          
          If nBusinessDays = nDays Then
          
             bLoop = False
             
          End If
    
    Wend
    
    Function = nSerial                               
                                   
End Function
' ========================================================================================
' Given a Gregorian Month and Year, calculate the first business day of the month
' Another use is to use January as the month and find out the first
' business day of a year
' ========================================================================================
Private Function cCalendar.BusinessMonthBegin (ByVal nMonth as Long, _
                                               ByVal nYear as Long) as LongInt

Dim bSearch               as BOOLEAN
Dim nFirstOfMonth      as LongInt
Dim nEndOfMonth           as LongInt
Dim nSearchDate           as LongInt
Dim uGreg                 as GREGORIAN_DATE
Dim nDay                  as LongInt

' Restrict search to month specified, return calendar first of month
' if no business days could be found

    uGreg.Month = nMonth
    uGreg.Day = 1
    uGreg.Year = nYear
    uGreg.Hour = 0
    uGreg.Minute = 0
    uGreg.Second = 0
    uGreg.Millisecond = 0

    nFirstOfMonth = SerialFromGregorian(uGreg)

' Calendar end of month

    uGreg.Day = GregorianDaysInMonth(nMonth,nYear)

    nEndOfMonth = SerialFromGregorian(uGreg)

' Start search on last day of previous month

    nSearchDate = nFirstOfMonth - cCalendarClass.ONE_DAY

    bSearch = False

    Do While bSearch = False

        nSearchDate = nSearchDate + cCalendarClass.ONE_DAY

        If nSearchDate > nEndOfMonth Then

           nSearchDate = nFirstOfMonth
           bSearch = True

        Else

           bSearch = IsBusinessDay(nSearchDate)
          
        End If

    Loop

    Function = nSearchDate

End Function
' ========================================================================================
' Given a Gregorian Month and Year, calculate the last business day of the month
' Another use is to use December as the month and find out the last
' business day of a year
' ========================================================================================
Private Function cCalendar.BusinessMonthEnd (ByVal nMonth as Long, _
                                             ByVal nYear as Long) as LongInt

Dim bSearch               as BOOLEAN
Dim nFirstOfMonth         as LongInt
Dim nEndOfMonth           as LongInt
Dim nSearchDate           as LongInt
Dim uGreg                 as GREGORIAN_DATE
Dim nDay                  as LongInt

' Restrict search to month specified, return calendar end of month
' if no business days could be found

    uGreg.Month = nMonth
    uGreg.Day = 1
    uGreg.Year = nYear
    uGreg.Hour = 0
    uGreg.Minute = 0
    uGreg.Second = 0
    uGreg.Millisecond = 0

    nFirstOfMonth = SerialFromGregorian(uGreg)

' Calendar end of month

    uGreg.Day = GregorianDaysInMonth(nMonth,nYear)

    nEndOfMonth = SerialFromGregorian(uGreg)

' Start search on first of next month

    nSearchDate = nEndOfMonth + cCalendarClass.ONE_DAY

    bSearch = False

    Do While bSearch = False

        nSearchDate = nSearchDate - cCalendarClass.ONE_DAY

        If nSearchDate < nFirstOfMonth Then

           nSearchDate = nFirstOfMonth
           bSearch = True

        Else

           bSearch = IsBusinessDay(nSearchDate)
           
        End If

    Loop

    Function = nSearchDate

End Function
' ========================================================================================
' Calculate the number of business days between two dates
' ========================================================================================
Private Function cCalendar.BusinessDaysBetween (ByVal nStart as LongInt, _
                                                ByVal nEnd as LongInt) as Long

' The starting/ending dates are not included in the count
' To get the business days in a month, start with the last day
' of the previous month and end with the first day of the following month

Dim nDays                  as Long
Dim nCalendarDays          as Long
Dim nLoop                  as Long

    nDays = 0

' Ensure that the starting date is less than the ending date

    If nStart > nEnd Then

        Swap nStart, nEnd

    End If

    nDays = cmFloor((nEnd - nStart) / cCalendarClass.ONE_DAY) - 1

    Select Case nDays

    Case Is > 1

        nCalendarDays = nDays - 1
        
        For nLoop = 1 To nCalendarDays

            nStart = nStart + cCalendarClass.ONE_DAY

            If IsBusinessDay(nStart) = False Then

                nDays = nDays - 1

            End If

        Next

    End Select

    Function = nDays

End Function
' ========================================================================================
' Calculate the number of business days between two dates
' ========================================================================================
Private Function cCalendar.NonBusinessDaysBetween (ByVal nStart as LongInt, _
                                                   ByVal nEnd as LongInt) as Long

' The starting/ending dates are not included in the count
' To get the non business days in a month, start with the last day
' of the previous month and end with the first day of the following month

Dim nDays                  as Long
Dim nCalendarDays          as Long
Dim nLoop                  as Long

    nDays = 0

' Ensure that the starting date is less than the ending date

    If nStart > nEnd Then

        Swap nStart, nEnd

    End If

    nDays = cmFloor((nEnd - nStart) / cCalendarClass.ONE_DAY) - 1

    Select Case nDays

    Case Is > 1

        nCalendarDays = nDays - 1
        nDays = 0
        
        For nLoop = 1 To nCalendarDays

            nStart = nStart + cCalendarClass.ONE_DAY

            If IsBusinessDay(nStart) = False Then

                nDays = nDays + 1

            End If

        Next

    End Select

    Function = nDays

End Function
' ========================================================================================
' Check if a serial date is a business day
' ========================================================================================
Private Function cCalendar.IsBusinessDay (ByVal nSerial as LongInt) as BOOLEAN

Dim nDays           as Long
Dim nWeekday        as Short
Dim nLowerIndex     as Long
Dim nHigherIndex    as Long
Dim nMidIndex       as Long
Dim bLoop           as BOOLEAN
Dim bFound          as BOOLEAN  

    nDays = cmFloor(nSerial / cCalendarClass.ONE_DAY)
    
' Check if this is a saved non business day

    If This.iUBound >= 0 Then
    
       If nDays < This.arNonBusinessDays(0) OrElse nDays > This.arNonBusinessDays(This.iUBound) Then
       
' If out of range, then only the weekday rules are applied

' This can happen when provided year of date has no rules defined.
       
          nWeekday = cmGregorianWeekday(nDays)
          Function = This.arBusinessWeekdays(nWeekday)
          Exit Function
          
       Else

' Perform bisection search of saved dates

          nLowerIndex = 0
          nHigherIndex = This.iUBound
          nMidIndex = cmFloor(This.iUBound / 2)
          bLoop = True
          bFound = False

          While bLoop = True
          
             If nDays = This.arNonBusinessDays(nMidIndex) Then
             
                bFound = True
                bLoop = False
                
            Else
            
' Is search completed?

               If nMidIndex = nLowerIndex OrElse nMidIndex = nHigherIndex Then
               
                  bLoop = False
                  
               Else
               
                  If nDays < This.arNonBusinessDays(nMidIndex) Then
                  
                     nHigherIndex = nMidIndex
                     
                  Else
                  
                     nLowerIndex = nMidIndex
                     
                  End If
                  
                  nMidIndex = cmFloor((nLowerIndex + nHigherIndex) / 2)

               End If                    
          
            End If
          
          Wend
          
          If bFound = True Then
          
             Function = False
             Exit Function
             
          End If        
            
       End If

    End If
    
' Check for day of week
    
    nWeekday = cmGregorianWeekDay(nDays)
    
    Function = This.arBusinessWeekdays(nWeekday)
        
End Function
' ========================================================================================
' Retrieve saved non business dates
' ========================================================================================
Private Sub cCalendar.GetSavedNonBusinessDates (arDates() as LongInt)

Dim iIndex  as Long

    Erase arDates
    
    If This.iUBound >=0 Then
    
       ReDim arDates(0 To iUBound) 
    
       For iIndex = 0 To This.iUBound
       
          arDates(iIndex) = This.arNonBusinessDays(iIndex) * cCalendarClass.ONE_DAY        
           
       Next
    
    End If

End Sub
' ========================================================================================
' Flag a weekday as a working day
' ========================================================================================
Private Sub cCalendar.SetBusinessWeekday (ByVal nWeekday as Short, _
                                          ByVal bWorkday as BOOLEAN)

    Select Case nWeekday
    
        Case cCalendarClass.SUNDAY To cCalendarClass.SATURDAY
 
           This.arBusinessWeekdays(nWeekday) = bWorkday
        
    End Select

End Sub
' ========================================================================================
' Retrieve weekday working day status
' ========================================================================================
Property cCalendar.GetBusinessWeekday (ByVal nWeekday as Short) as BOOLEAN
  
    Select Case nWeekday
    
        Case cCalendarClass.SUNDAY To cCalendarClass.SATURDAY
 
           Property = This.arBusinessWeekdays(nWeekday)
           
        Case Else
        
           Property = False           
        
    End Select

End Property
' ========================================================================================
' Retrieve maximum number of non business days supported
' ========================================================================================
Property cCalendar.NonBusinessDatesLimit() as Long

    Property = UBound(This.arNonBusinessDays) + 1

End Property
' ========================================================================================
' Retrieve number of non business days saved
' ========================================================================================
Property cCalendar.NonBusinessDatesSaved() as Long

    Property = This.iUBound + 1

End Property
' ========================================================================================
' Erase Saved Non Business Days
' ========================================================================================
Private Sub cCalendar.EraseNonBusinessDates()

    This.iUBound = -1

End Sub

' ########################################################################################
' Date Calculator
' ########################################################################################

' ========================================================================================
' Date Calculations
' ========================================================================================
Private Sub cCalendar.DateCalculation (arRules() as DATE_CALCULATION)

' Results, in Days format is returned in arDays to handle non-gregorian
' calendars where some holidays may occur more than once during the
' Jan-Dec gregorian year

Dim nIndex            as Long
Dim nCalcDays         as Long
Dim nCalcDays1        as Long

    If UBound(arRules) >= 0 Then                 ' Check if any rules are available

        For nIndex = 0 To UBound(arRules)
        
            arRules(nIndex).MaxNonBusinessDates = False

            Select Case arRules(nIndex).RuleClass

            Case cCalendarClass.GREGORIAN_RULES

                cmGregorianDateCalculation ( _
                                    arRules(nIndex).Month, _
                                    arRules(nIndex).Day, _
                                    arRules(nIndex).Year, _
                                    arRules(nIndex).Weekday, _
                                    arRules(nIndex).Rule, _
                                    nCalcDays)

' Save initial date before applying observation rules

                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY

' Apply Observed/Year Rules

                arRules(nIndex).Observed = cmDateCalculationObserved ( _
                                                       arRules(nIndex).ThursdayRule, _
                                                       arRules(nIndex).FridayRule, _
                                                       arRules(nIndex).SaturdayRule, _
                                                       arRules(nIndex).SundayRule, _
                                                       arRules(nIndex).MondayRule, _
                                                       arRules(nIndex).YearRule, _
                                                       nCalcDays, _
                                                       arRules(nIndex).Year, _
                                                       arRules(nIndex).ObservedDays2)
                                
' Convert Observed Days to Serial Date

                arRules(nIndex).ObservedDays2 = arRules(nIndex).ObservedDays2 * cCalendarClass.ONE_DAY
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))                                

            Case cCalendarClass.CHRISTIANEASTER_RULES, cCalendarClass.ORTHODOXEASTER_RULES

                arRules(nIndex).Month = 0
                arRules(nIndex).Day = 0
                arRules(nIndex).Observed = True
                arRules(nIndex).YearRule = cCalendarClass.ALL_YEARS
                arRules(nIndex).Weekday = cCalendarClass.ALL_WEEKDAYS
                arRules(nIndex).SaturdayRule = cCalendarClass.NO_SATURDAY_RULE
                arRules(nIndex).SundayRule = cCalendarClass.NO_SUNDAY_RULE
                cmEasterCalculation ( _
                                arRules(nIndex).RuleClass, _
                                arRules(nIndex).Rule, _
                                arRules(nIndex).Year, _
                                nCalcDays)

' No Observed Rules are used

                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = arRules(nIndex).ObservedDays1
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))

            Case cCalendarClass.HEBREW_RULES

                arRules(nIndex).ObservedDays1 = 0
                arRules(nIndex).ObservedDays2 = 0
                arRules(nIndex).Observed = cmHebrewDateCalculation ( _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    arRules(nIndex).Weekday, _
                                                    arRules(nIndex).Rule, _
                                                    nCalcDays)
                                                    
' Save initial date before applying observation rules

                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY

                If arRules(nIndex).Observed = True Then
                
' Apply Observed/Year Rules

                   arRules(nIndex).Observed = cmDateCalculationObserved ( _
                                                       arRules(nIndex).ThursdayRule, _
                                                       arRules(nIndex).FridayRule, _
                                                       arRules(nIndex).SaturdayRule, _
                                                       arRules(nIndex).SundayRule, _
                                                       arRules(nIndex).MondayRule, _
                                                       arRules(nIndex).YearRule, _
                                                       nCalcDays, _
                                                       arRules(nIndex).Year, _
                                                       arRules(nIndex).ObservedDays2)


' Convert Observed Days to Serial Date

                   arRules(nIndex).ObservedDays2 = arRules(nIndex).ObservedDays2 * cCalendarClass.ONE_DAY

                End If
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))

            Case cCalendarClass.CHINESE_RULES,cCalendarClass.KOREAN_RULES,cCalendarClass.VIETNAMESE_RULES,cCalendarClass.JAPANESE_RULES

                arRules(nIndex).ObservedDays1 = 0
                arRules(nIndex).ObservedDays2 = 0
                arRules(nIndex).Observed = cmChineseDateCalculation ( _
                                                    arRules(nIndex).RuleClass, _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    arRules(nIndex).Rule, _
                                                    nCalcDays)
 
' No Observed Rules are used 
                                                    
                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = arRules(nIndex).ObservedDays1

' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))

            Case cCalendarClass.PERSIAN_RULES

                arRules(nIndex).ObservedDays1 = 0
                arRules(nIndex).ObservedDays2 = 0
                arRules(nIndex).Observed = cmPersianDateCalculation ( _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    nCalcDays)
                                                    
' No Observed Rules are used 
                                                    
                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = arRules(nIndex).ObservedDays1
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))
                
            Case cCalendarClass.BAHAI_RULES

                arRules(nIndex).Observed = cmBahaiDateCalculation ( _
                                                    arRules(nIndex).Rule, _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    nCalcDays)
                                                    
' No Observed Rules are used 
                                                    
                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = nCalcDays1 * cCalendarClass.ONE_DAY
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))
                
            Case cCalendarClass.TIBETAN_RULES

                arRules(nIndex).Observed = cmTibetanDateCalculation ( _
                                                    arRules(nIndex).Rule, _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    nCalcDays)
                                                    
' No Observed Rules are used 
                                                    
                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = nCalcDays1 * cCalendarClass.ONE_DAY
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))

            Case cCalendarClass.HINDU_SOLAR_RULES

                arRules(nIndex).ObservedDays1 = 0
                arRules(nIndex).ObservedDays2 = 0
                arRules(nIndex).Observed = cmHinduSolarDateCalculation ( _
                                                    arRules(nIndex).Rule, _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    nCalcDays)
                                                    
' No Observed Rules are used 
                                                    
                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = arRules(nIndex).ObservedDays1
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))

            Case cCalendarClass.HINDU_LUNAR_RULES

                arRules(nIndex).Observed = cmHinduLunarDateCalculation ( _
                                                    arRules(nIndex).Rule, _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    nCalcDays, _
                                                    nCalcDays1)
                                                    
' No Observed Rules are used 
                                                    
                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = nCalcDays1 * cCalendarClass.ONE_DAY
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))

            Case cCalendarClass.COPTIC_RULES

                arRules(nIndex).ObservedDays1 = 0
                arRules(nIndex).ObservedDays2 = 0
                arRules(nIndex).Observed = cmCopticDateCalculation ( _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    nCalcDays)

' No Observed Rules are used 
                                                    
                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = arRules(nIndex).ObservedDays1
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))   

            Case cCalendarClass.ETHIOPIC_RULES

                arRules(nIndex).ObservedDays1 = 0
                arRules(nIndex).ObservedDays2 = 0
                arRules(nIndex).Observed = cmEthiopicDateCalculation ( _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    nCalcDays)

' No Observed Rules are used 
                                                    
                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = arRules(nIndex).ObservedDays1
        
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))
                
            Case cCalendarClass.ISLAMIC_RULES

                arRules(nIndex).Observed = cmIslamicDateCalculation ( _
                                                    arRules(nIndex).Rule, _
                                                    arRules(nIndex).Month, _
                                                    arRules(nIndex).Day, _
                                                    arRules(nIndex).Year, _
                                                    nCalcDays, _
                                                    nCalcDays1)
                                                    
' No Observed Rules are used 
                                                    
                arRules(nIndex).ObservedDays1 = nCalcDays * cCalendarClass.ONE_DAY
                arRules(nIndex).ObservedDays2 = nCalcDays1 * cCalendarClass.ONE_DAY
                
' Save date for business day calculations

                cmSaveDateCalculation (arRules(nIndex))
                
            Case Else
            
                arRules(nIndex).Observed = False
                arRules(nIndex).ObservedDays1 = 0
                arRules(nIndex).ObservedDays2 = 0

            End Select

' Apply Week Begin Rule

            nCalcDays = arRules(nIndex).ObservedDays1 / cCalendarClass.ONE_DAY
            nCalcDays1 = arRules(nIndex).ObservedDays2 / cCalendarClass.ONE_DAY
            arRules(nIndex).ObservedDaysBegin1 = cmWeekDayOnOrBefore(arRules(nIndex).WeekRuleWeekday,nCalcDays)
            arRules(nIndex).ObservedDaysBegin2 = cmWeekDayOnOrBefore(arRules(nIndex).WeekRuleWeekday,nCalcDays1)

        Next

    End If

End Sub

' ########################################################################################
' Gregorian Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid gregorian date representation
' ========================================================================================
Private Function cCalendar.ValidGregorianDate (ByRef uDate as GREGORIAN_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as GREGORIAN_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmGregorianFromDays(cmDaysFromGregorian(uDate.Month,uDate.Day,uDate.Year), _
                        uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Gregorian Date from Serial
' ========================================================================================
Private Sub cCalendar.GregorianFromSerial (ByVal nSerial as LongInt, _
                                           ByRef uDate as GREGORIAN_DATE)

Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmGregorianFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
    uDate.Weekday = cmGregorianWeekDay(nDays)
                  
End Sub
' ========================================================================================
' Serial Date from Gregorian
' ========================================================================================
Private Function cCalendar.SerialFromGregorian (ByRef uDate as GREGORIAN_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromGregorian(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function
' ========================================================================================
' Days remaining in Gregorian year
' ========================================================================================
Private Function cCalendar.GregorianDaysRemaining (ByVal nMonth as Short, _
                                                   ByVal nDay as Short, _
                                                   ByVal nYear as Long) as Short

Dim nEndMonth      as Short
Dim nEndDay        as Short
Dim nEndYear       as Long

' Get last day of year

    cmGregorianFromDays(cmGregorianYearEnd(nYear), _
                        nEndMonth, _
                        nEndDay, _
                        nEndYear)

   Function = cmGregorianDateDifference(nMonth, _
                                        nDay, _
                                        nYear, _
                                        nEndMonth, _
                                        nEndDay, _
                                        nEndYear)

End Function
' ========================================================================================
' Calculate the Gregorian ordinal day of the year (1-365 or 366 if a leap year)
' ========================================================================================
Private Function cCalendar.GregorianDayNumber (ByVal nMonth as Short, _
                                               ByVal nDay as Short, _
                                               ByVal nYear as Long) as Short

Dim nPriorEndMonth      as Short
Dim nPriorEndDay        as Short
Dim nPriorEndYear       as Long

' Get last day of prior year

    cmGregorianFromDays(cmGregorianYearEnd(nYear - 1), _
                        nPriorEndMonth, _
                        nPriorEndDay, _
                        nPriorEndYear)

    Function = cmGregorianDateDifference(nPriorEndMonth, _
                                         nPriorEndDay, _
                                         nPriorEndYear, _
                                         nMonth, _
                                         nDay, _
                                         nYear)

End Function
' ========================================================================================
' Calculate the number of days in a month occurring in a Year
' ========================================================================================
Private Function cCalendar.GregorianDaysInMonth (ByVal nMonth as Short, _
                                                 ByVal nYear as Long) as Short

' Since we need to use the first day of the next month, if
' nMonth is December, we have to roll to January of the next year

    Function = cmGregorianDateDifference(nMonth, _
                                       1, _
                                       nYear, _
                                       IIf(nMonth = cCalendarClass.DECEMBER,cCalendarClass.JANUARY,nMonth + 1), _
                                       1, _
                                       IIf(nMonth = cCalendarClass.DECEMBER,nYear + 1,nYear))

End Function
' ========================================================================================
' Given a calendar month, return the fiscal quarter it belongs to
' ========================================================================================
Private Function cCalendar.GregorianFiscalQuarter (ByVal nCalendarMonth as Short, _
                                                   ByVal nFiscalYearBeginsMonth as Short) as Short

    Function = cmCeiling(GregorianFiscalMonth(nCalendarMonth,nFiscalYearBeginsMonth) / 3)

End Function
' ========================================================================================
' Given a calendar month, return the fiscal month
' ========================================================================================
Private Function cCalendar.GregorianFiscalMonth (ByVal nCalendarMonth as Short, _
                                                 ByVal nFiscalYearBeginsMonth as Short) as Short

    Function = nCalendarMonth _
             + IIf(nCalendarMonth < nFiscalYearBeginsMonth,12,0) _
             - nFiscalYearBeginsMonth _
             + 1

End Function
' ========================================================================================
' Given a month and year, return the fiscal years for months Jan-Dec
' ========================================================================================
Private Sub cCalendar.GregorianFiscalMonthYears (ByVal nCalendarYear as Long, _
                                                 ByVal nFiscalYearBeginsMonth as Short, _
                                                 ByVal bUseEndingMonth as BOOLEAN,_
                                                 arFiscalYears() as Long)

Dim nMonth        as Short

    ReDim arFiscalYears(11)

    For nMonth = 0 To 11

        arFiscalYears(nMonth) = GregorianFiscalYear(_
                                nMonth + 1, _
                                nCalendarYear, _
                                nFiscalYearBeginsMonth, _
                                bUseEndingMonth)

    Next

End Sub
' ========================================================================================
' Given a month and year, return the fiscal year
' ========================================================================================
Private Function cCalendar.GregorianFiscalYear (ByVal nCalendarMonth as Short, _
                                                ByVal nCalendarYear as Long, _
                                                ByVal nFiscalYearBeginsMonth as Short, _
                                                ByVal bUseEndingMonth as BOOLEAN) as Long

' bUseEndingMonth

'     TRUE  = Use the year of the ending month
'     FALSE = Use the year of the beginning month

    Function = IIf(nCalendarMonth < nFiscalYearBeginsMonth,nCalendarYear - 1,nCalendarYear)

End Function

' ########################################################################################
' Chinese Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid chinese date representation
' ========================================================================================
Private Function cCalendar.ValidChineseDate (ByRef uDate as CHINESE_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as CHINESE_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

' Validate country code

    Select Case uDate.Country
    
        Case cCalendarClass.CHINESE, cCalendarClass.VIETNAMESE, cCalendarClass.KOREAN, cCalendarClass.JAPANESE
        
        Case Else
        
            uDate.Country = cCalendarClass.CHINESE
            
    End Select

    cmChineseFromDays(cmDaysFromChinese(uDate.Cycle,uDate.Year,uDate.Month,uDate.LeapMonth,uDate.Day,uDate.Country), _
                      uWorkDate.Cycle,uWorkDate.Year,uWorkDate.Month,uWorkDate.LeapMonth,uWorkDate.Day,uWorkDate.Country)

    If uWorkDate.Cycle <> uDate.Cycle Then

       Function = cCalendarClass.INVALID_CYCLE

    Else

       If uWorkDate.Year <> uDate.Year Then

          Function = cCalendarClass.INVALID_YEAR

       Else

          If uWorkDate.Month <> uDate.Month Then

             Function = cCalendarClass.INVALID_MONTH

          Else

             If uWorkDate.LeapMonth <> uDate.LeapMonth Then

                Function = cCalendarClass.INVALID_LEAP_MONTH

             Else

                If uWorkDate.Day <> uDate.Day Then

                   Function = cCalendarClass.INVALID_DAY

                Else 

                   Function = cCalendarClass.VALID_DATE

                End If

             End If

          End If

       End If

    End If

End Function
' ========================================================================================
' Chinese Date from Serial
' ========================================================================================
Private Sub cCalendar.ChineseFromSerial (ByVal nSerial as LongInt, _
                                         ByRef uDate as CHINESE_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long

' Validate country code

    Select Case uDate.Country
    
        Case cCalendarClass.CHINESE, cCalendarClass.VIETNAMESE, cCalendarClass.KOREAN, cCalendarClass.JAPANESE
        
        Case Else
        
            uDate.Country = cCalendarClass.CHINESE
            
    End Select 
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmChineseFromDays(nDays,uDate.Cycle,uDate.Year,uDate.Month,uDate.LeapMonth,uDate.Day,uDate.Country)
    uDate.YearAnimal = cmChineseYearName(uDate.Year)
    uDate.MonthAnimal = cmChineseMonthName(uDate.Month,uDate.Year)
    uDate.YearAugury = cmChineseYearMarriageAuguries(uDate.Cycle,uDate.Year,uDate.Country)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub                         
' ========================================================================================
' Serial Date from Chinese
' ========================================================================================
Private Function cCalendar.SerialFromChinese (ByRef uDate as CHINESE_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

' Validate country code

    Select Case uDate.Country
    
        Case cCalendarClass.CHINESE, cCalendarClass.VIETNAMESE, cCalendarClass.KOREAN, cCalendarClass.JAPANESE
        
        Case Else
        
            uDate.Country = cCalendarClass.CHINESE
            
    End Select

    nSerialDays = cmDaysFromChinese(uDate.Cycle,uDate.Year,uDate.Month,uDate.LeapMonth,uDate.Day,uDate.Country)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Hebrew Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid hebrew date representation
' ========================================================================================
Private Function cCalendar.ValidHebrewDate (ByRef uDate as HEBREW_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as HEBREW_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmHebrewFromDays(cmDaysFromHebrew(uDate.Month,uDate.Day,uDate.Year), _
                     uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Hebrew Date from Serial
' ========================================================================================
Private Sub cCalendar.HebrewFromSerial (ByVal nSerial as LongInt, _
                                        ByRef uDate as HEBREW_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmHebrewFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
    uDate.SabbaticalYear = cmHebrewSabbaticalYear(uDate.Year)
                       
End Sub                         
' ========================================================================================
' Serial Date from Hebrew
' ========================================================================================
Private Function cCalendar.SerialFromHebrew (ByRef uDate as HEBREW_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromHebrew(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function
' ========================================================================================
' Hebrew Birthday in a given Hebrew Year
' ========================================================================================
Private Sub cCalendar.HebrewBirthDay (ByRef uBirth as HEBREW_DATE, _
                                      ByVal nForHebrewYear as Long, _
                                      ByRef uBirthDay as HEBREW_DATE)
                   
Dim nDays   as Long

    nDays = cmHebrewBirthDay(uBirth.Month,uBirth.Day,uBirth.Year,nForHebrewYear)
                   
    cmHebrewFromDays(nDays,uBirthDay.Month,uBirthDay.Day,uBirthDay.Year)
    uBirthDay.Hour = 0
    uBirthDay.Minute = 0
    uBirthDay.Second = 0
    uBirthDay.Millisecond = 0               
                   
End Sub
' ========================================================================================
' Hebrew Birthday in a given Gregorian Year
' ========================================================================================
Private Sub cCalendar.HebrewBirthDayInGregorian (ByRef uBirth as HEBREW_DATE, _
                                                 ByVal nForGregorianYear as Long, _
                                                 ByRef uBirthDay1 as GREGORIAN_DATE, _
                                                 ByRef uBirthDay2 as GREGORIAN_DATE)
                              
' A Hebrew birthday may occur twice in a Gregorian Year
' If there is a second birthday, uBirthDay2.Month will be non-zero

Dim nJan1           as Long
Dim nDec31          as Long
Dim nHebrewMonth    as Short
Dim nHebrewDay      as Short
Dim nHebrewYear     as Long
Dim nBirthDay1      as Long
Dim nBIrthDay2      as Long

    uBirthDay2.Month = 0
    uBirthDay2.Day = 0
    uBirthDay2.Year = 0
    uBirthDay2.Hour = 0
    uBirthDay2.Minute = 0
    uBirthDay2.Second = 0
    uBirthDay2.Millisecond = 0
    uBirthDay2.LeapYear = 0
    uBirthDay2.Weekday = 0
    uBirthDay1.Hour = 0
    uBirthDay1.Minute = 0
    uBirthDay1.Second = 0
    uBirthDay1.Millisecond = 0
    
    cmGregorianYearRange(nForGregorianYear,nJan1,nDec31)
    cmHebrewFromDays(nJan1,nHebrewMonth,nHebrewDay,nHebrewYear)
    nBirthDay1 = cmHebrewBirthDay(uBirth.Month,uBirth.Day,uBirth.Year,nHebrewYear)
    nBirthDay2 = cmHebrewBirthDay(uBirth.Month,uBirth.Day,uBirth.Year,nHebrewYear + 1)
    
' Now see which birthdays fall in the requested Gregorian Year, one always will

    If nBirthDay1 >= nJan1 AndAlso nBirthDay1 <= nDec31 Then
    
       cmGregorianFromDays(nBirthDay1,uBirthDay1.Month,uBirthDay1.Day,uBirthDay1.Year)
       uBirthDay1.Weekday = cmGregorianWeekDay(nBirthDay1)
       uBirthDay1.LeapYear = cmGregorianLeapYear(uBirthDay1.Year)
       
       If nBirthDay2 >= nJan1 AndAlso nBirthDay2 <= nDec31 Then

          cmGregorianFromDays(nBirthDay2,uBirthDay2.Month,uBirthDay2.Day,uBirthDay2.Year)
          uBirthDay2.Weekday = cmGregorianWeekDay(nBirthDay2)
          uBirthDay2.LeapYear = cmGregorianLeapYear(uBirthDay2.Year)
          
       End If       
       
    Else
    
       If nBirthDay2 >= nJan1 AndAlso nBirthDay2 <= nDec31 Then

          cmGregorianFromDays(nBirthDay1,uBirthDay1.Month,uBirthDay1.Day,uBirthDay1.Year)
          uBirthDay1.Weekday = cmGregorianWeekDay(nBirthDay1)
          uBirthDay1.LeapYear = cmGregorianLeapYear(uBirthDay1.Year)
          
       End If    
    
    End If    

End Sub

' ########################################################################################
' Islamic Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid islamic date representation
' ========================================================================================
Private Function cCalendar.ValidIslamicDate (ByRef uDate as ISLAMIC_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as ISLAMIC_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmIslamicFromDays(cmDaysFromIslamic(uDate.Month,uDate.Day,uDate.Year), _
                      uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Islamic Date from Serial
' ========================================================================================
Private Sub cCalendar.IslamicFromSerial (ByVal nSerial as LongInt, _
                                         ByRef uDate as ISLAMIC_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmIslamicFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub                         
' ========================================================================================
' Serial Date from Islamic
' ========================================================================================
Private Function cCalendar.SerialFromIslamic(ByRef uDate as ISLAMIC_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromIslamic(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Persian Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid persian date representation
' ========================================================================================
Private Function cCalendar.ValidPersianDate (ByRef uDate as PERSIAN_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as PERSIAN_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmPersianFromDays(cmDaysFromPersian(uDate.Month,uDate.Day,uDate.Year), _
                      uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Persian Date from Serial
' ========================================================================================
Private Sub cCalendar.PersianFromSerial(ByVal nSerial as LongInt, _
                                        ByRef uDate as PERSIAN_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmPersianFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub                         
' ========================================================================================
' Serial Date from Persian
' ========================================================================================
Private Function cCalendar.SerialFromPersian(ByRef uDate as PERSIAN_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromPersian(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Hindu Calendars
' ########################################################################################

' ========================================================================================
' Validate the date as a valid hindu lunar date representation
' ========================================================================================
Private Function cCalendar.ValidHinduLunarDate (ByRef uDate as HINDU_LUNAR_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as HINDU_LUNAR_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmHinduLunarFromDays(cmDaysFromHinduLunar(uDate.Month,uDate.LeapMonth,uDate.Day,uDate.LeapDay,uDate.Year), _
                         uWorkDate.Month,uWorkDate.LeapMonth,uWorkDate.Day,uWorkDate.LeapDay,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.LeapDay <> uDate.LeapDay Then

          Function = cCalendarClass.INVALID_LEAP_DAY

       Else

          If uWorkDate.Month <> uDate.Month Then

             Function = cCalendarClass.INVALID_MONTH

          Else

             If uWorkDate.LeapMonth <> uDate.LeapMonth Then

                Function = cCalendarClass.INVALID_LEAP_MONTH

             Else 

                Function = cCalendarClass.VALID_DATE

             End If

          End If

       End If

    End If

End Function
' ========================================================================================
' Hindu Lunar Date from Serial
' ========================================================================================
Private Sub cCalendar.HinduLunarFromSerial (ByVal nSerial as LongInt, _
                                            ByRef uDate as HINDU_LUNAR_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmHinduLunarFromDays(nDays,uDate.Month,uDate.LeapMonth,uDate.Day,uDate.LeapDay,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub                         
' ========================================================================================
' Serial Date from Hindu Lunar
' ========================================================================================
Private Function cCalendar.SerialFromHinduLunar(ByRef uDate as HINDU_LUNAR_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromHinduLunar(uDate.Month,uDate.LeapMonth,uDate.Day,uDate.LeapDay,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function
' ========================================================================================
' Validate the date as a valid hindu solar date representation
' ========================================================================================
Private Function cCalendar.ValidHinduSolarDate (ByRef uDate as HINDU_SOLAR_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as HINDU_SOLAR_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmHinduSolarFromDays(cmDaysFromHinduSolar(uDate.Month,uDate.Day,uDate.Year), _
                         uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Hindu Solar Date from Serial
' ========================================================================================
Private Sub cCalendar.HinduSolarFromSerial(ByVal nSerial as LongInt, _
                                           ByRef uDate as HINDU_SOLAR_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmHinduSolarFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub                         
' ========================================================================================
' Serial Date from Hindu Solar
' ========================================================================================
Private Function cCalendar.SerialFromHinduSolar(ByRef uDate as HINDU_SOLAR_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromHinduSolar(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Coptic Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid coptic date representation
' ========================================================================================
Private Function cCalendar.ValidCopticDate (ByRef uDate as COPTIC_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as COPTIC_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmCopticFromDays(cmDaysFromCoptic(uDate.Month,uDate.Day,uDate.Year), _
                     uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Coptic Date from Serial
' ========================================================================================
Private Sub cCalendar.CopticFromSerial (ByVal nSerial as LongInt, _
                                        ByRef uDate as COPTIC_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmCopticFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub
' ========================================================================================
' Serial Date from Coptic
' ========================================================================================
Private Function cCalendar.SerialFromCoptic(ByRef uDate as COPTIC_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromCoptic(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Ethiopic Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid ethiopic date representation
' ========================================================================================
Private Function cCalendar.ValidEthiopicDate (ByRef uDate as ETHIOPIC_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as ETHIOPIC_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmEthiopicFromDays(cmDaysFromEthiopic(uDate.Month,uDate.Day,uDate.Year), _
                       uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Ethiopic Date from Serial
' ========================================================================================
Private Sub cCalendar.EthiopicFromSerial (ByVal nSerial as LongInt, _
                                          ByRef uDate as ETHIOPIC_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmEthiopicFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub
' ========================================================================================
' Serial Date from Ethiopic
' ========================================================================================
Private Function cCalendar.SerialFromEthiopic (ByRef uDate as ETHIOPIC_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromEthiopic(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' ISO Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid iso date representation
' ========================================================================================
Private Function cCalendar.ValidISODate (ByRef uDate as ISO_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as ISO_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmISOFromDays(cmDaysFromISO(uDate.Week,uDate.Day,uDate.Year), _
                  uWorkDate.Week,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Week <> uDate.Week Then

          Function = cCalendarClass.INVALID_WEEK

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' ISO Date from Serial
' ========================================================================================
Private Sub cCalendar.ISOFromSerial (ByVal nSerial as LongInt, _
                                     ByRef uDate as ISO_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmISOFromDays(nDays,uDate.Week,uDate.Day,uDate.Year)
    uDate.LongYear = cmISOLongYear(uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub
' ========================================================================================
' Serial Date from ISO
' ========================================================================================
Private Function cCalendar.SerialFromISO (ByRef uDate as ISO_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromISO(uDate.Week,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Julian Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid julian date representation
' ========================================================================================
Private Function cCalendar.ValidJulianDate (ByRef uDate as JULIAN_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as JULIAN_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmJulianFromDays(cmDaysFromJulian(uDate.Month,uDate.Day,uDate.Year), _
                     uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Julian Date from Serial
' ========================================================================================
Private Sub cCalendar.JulianFromSerial (ByVal nSerial as LongInt, _
                                        ByRef uDate as JULIAN_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmJulianFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
    uDate.Weekday = cmGregorianWeekDay(nDays)
    uDate.LeapYear = cmJulianLeapYear(uDate.Year)
                        
End Sub
' ========================================================================================
' Serial Date from Julian
' ========================================================================================
Private Function cCalendar.SerialFromJulian (ByRef uDate as JULIAN_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromJulian(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Roman Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid roman date representation
' ========================================================================================
Private Function cCalendar.ValidRomanDate (ByRef uDate as ROMAN_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as ROMAN_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmRomanFromDays(cmDaysFromRoman(uDate.Month,uDate.Year,uDate.Event,uDate.Count,uDate.Leap), _
                    uWorkDate.Month,uWorkDate.Year,uWorkDate.Event,uWorkDate.Count,uDate.Leap)

    If uWorkDate.Month <> uDate.Month Then

       Function = cCalendarClass.INVALID_MONTH

    Else

       If uWorkDate.Event <> uDate.Event Then

          Function = cCalendarClass.INVALID_EVENT

       Else

          If uWorkDate.Count <> uDate.Count Then

             Function = cCalendarClass.INVALID_COUNT

          Else
          
             If uWorkDate.Leap <> uDate.Leap Then
             
                Function = cCalendarClass.INVALID_LEAP_YEAR
                
             Else

                Function = cCalendarClass.VALID_DATE
                
             End If

          End If  

       End If

    End If

End Function
' ========================================================================================
' Roman Date from Serial
' ========================================================================================
Private Sub cCalendar.RomanFromSerial (ByVal nSerial as LongInt, _
                                       ByRef uDate as ROMAN_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmRomanFromDays(nDays,uDate.Month,uDate.Year,uDate.Event,uDate.Count,uDate.Leap)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub
' ========================================================================================
' Serial Date from Roman
' ========================================================================================
Private Function cCalendar.SerialFromRoman (ByRef uDate as ROMAN_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromRoman(uDate.Month,uDate.Year,uDate.Event,uDate.Count,uDate.Leap)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Egyptian Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid egyptian date representation
' ========================================================================================
Private Function cCalendar.ValidEgyptianDate (ByRef uDate as EGYPTIAN_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as EGYPTIAN_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmEgyptianFromDays(cmDaysFromEgyptian(uDate.Month,uDate.Day,uDate.Year), _
                       uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Egyptian Date from Serial
' ========================================================================================
Private Sub cCalendar.EgyptianFromSerial (ByVal nSerial as LongInt, _
                                          ByRef uDate as EGYPTIAN_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmEgyptianFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub
' ========================================================================================
' Serial Date from Egyptian
' ========================================================================================
Private Function cCalendar.SerialFromEgyptian (ByRef uDate as EGYPTIAN_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromEgyptian(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Armenian Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid armenian date representation
' ========================================================================================
Private Function cCalendar.ValidArmenianDate (ByRef uDate as ARMENIAN_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as ARMENIAN_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmArmenianFromDays(cmDaysFromArmenian(uDate.Month,uDate.Day,uDate.Year), _
                       uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.Month <> uDate.Month Then

          Function = cCalendarClass.INVALID_MONTH

       Else

          Function = cCalendarClass.VALID_DATE

       End If

    End If

End Function
' ========================================================================================
' Armenian Date from Serial
' ========================================================================================
Private Sub cCalendar.ArmenianFromSerial (ByVal nSerial as LongInt, _
                                          ByRef uDate as ARMENIAN_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmArmenianFromDays(nDays,uDate.Month,uDate.Day,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub
' ========================================================================================
' Serial Date from Armenian
' ========================================================================================
Private Function cCalendar.SerialFromArmenian (ByRef uDate as ARMENIAN_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromArmenian(uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Bahai Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid bahai date representation
' ========================================================================================
Private Function cCalendar.ValidBahaiDate (ByRef uDate as BAHAI_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as BAHAI_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmBahaiFromDays(cmDaysFromBahai(uDate.Major,uDate.Cycle,uDate.Month,uDate.Day,uDate.Year), _
                    uWorkDate.Major,uWorkDate.Cycle,uWorkDate.Month,uWorkDate.Day,uWorkDate.Year)

    If uWorkDate.Major <> uDate.Major Then

       Function = cCalendarClass.INVALID_MAJOR

    Else

       If uWorkDate.Cycle <> uDate.Cycle Then

          Function = cCalendarClass.INVALID_CYCLE

       Else

          If uWorkDate.Month <> uDate.Month Then

             Function = cCalendarClass.INVALID_MONTH

          Else

             If uWorkDate.Day <> uDate.Day Then

                Function = cCalendarClass.INVALID_DAY

             Else

                If uWorkDate.Year <> uDate.Year Then

                   Function = cCalendarClass.INVALID_YEAR

                Else  

                   Function = cCalendarClass.VALID_DATE

                End If

            End If

          End If  

       End If

    End If

End Function
' ========================================================================================
' Bahai Date from Serial
' ========================================================================================
Private Sub cCalendar.BahaiFromSerial (ByVal nSerial as LongInt, _
                                       ByRef uDate as BAHAI_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmBahaiFromDays(nDays,uDate.Major,uDate.Cycle,uDate.Month,uDate.Day,uDate.Year)
    uDate.Era = ((uDate.Major * 361) - 361) + ((uDate.Cycle - 1) * 19) +  uDate.Year
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub                         
' ========================================================================================
' Serial Date from Bahai
' ========================================================================================
Private Function cCalendar.SerialFromBahai (ByRef uDate as BAHAI_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromBahai(uDate.Major,uDate.Cycle,uDate.Month,uDate.Day,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Tibetan Calendar
' ########################################################################################

' ========================================================================================
' Validate the date as a valid tibetan date representation
' ========================================================================================
Private Function cCalendar.ValidTibetanDate (ByRef uDate as TIBETAN_DATE) as Short

' Return VALID_DATE if valid, otherwise a value showing which part of the date was invalid

Dim uWorkDate    as TIBETAN_DATE

' The basic validation process is to convert the date to days format and back
' and then compare the month and day

    cmTibetanFromDays(cmDaysFromTibetan(uDate.Month,uDate.LeapMonth,uDate.Day,uDate.LeapDay,uDate.Year), _
                      uWorkDate.Month,uWorkDate.LeapMonth,uWorkDate.Day,uWorkDate.LeapDay,uWorkDate.Year)

    If uWorkDate.Day <> uDate.Day Then

       Function = cCalendarClass.INVALID_DAY

    Else

       If uWorkDate.LeapDay <> uDate.LeapDay Then

          Function = cCalendarClass.INVALID_LEAP_DAY

       Else

          If uWorkDate.Month <> uDate.Month Then

             Function = cCalendarClass.INVALID_MONTH

          Else

             If uWorkDate.LeapMonth <> uDate.LeapMonth Then

                Function = cCalendarClass.INVALID_LEAP_MONTH

             Else 

                Function = cCalendarClass.VALID_DATE

             End If

          End If

       End If

    End If

End Function
' ========================================================================================
' Tibetan Date from Serial
' ========================================================================================
Private Sub cCalendar.TibetanFromSerial (ByVal nSerial as LongInt, _
                                         ByRef uDate as TIBETAN_DATE)
                        
Dim nDays       as Long
Dim nTime       as Long
                        
    cmSerialBreakApart(nSerial,nDays,nTime)
    cmTibetanFromDays(nDays,uDate.Month,uDate.LeapMonth,uDate.Day,uDate.LeapDay,uDate.Year)
    cmTimeFromSerial(nTime,uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                        
End Sub                         
' ========================================================================================
' Serial Date from Tibetan
' ========================================================================================
Private Function cCalendar.SerialFromTibetan (ByRef uDate as TIBETAN_DATE) as LongInt

Dim nSerialDays as LongInt
Dim nSerialTime as LongInt

    nSerialDays = cmDaysFromTibetan(uDate.Month,uDate.LeapMonth,uDate.Day,uDate.LeapDay,uDate.Year)
    nSerialTime = cmTimeToSerial(uDate.Hour,uDate.Minute,uDate.Second,uDate.Millisecond)
                                 
    Function = (Abs(nSerialDays) * cCalendarClass.ONE_DAY + nSerialTime) * IIf(nSerialDays < 0,-1,1)                             
                             
End Function

' ########################################################################################
' Time Calculations
' ########################################################################################

' ========================================================================================
' Serial Time Addition
' ========================================================================================
Private Function cCalendar.SerialAddTime (ByVal nSerial as LongInt, _
                                          ByRef uTime as TIME_DURATION) as LongInt
                       
' The time duration can be any combination of values
' in the range of the type supported. Adding 30 hours is
' equivalent to adding 1 day and 6 hours.

    Function = nSerial + TimeToSerial(uTime)
                       
End Function
' ========================================================================================
' Serial Time Subtract
' ========================================================================================
Private Function cCalendar.SerialSubtractTime (ByVal nSerial as LongInt, _
                                               ByRef uTime as TIME_DURATION) as LongInt
                       
' The time duration can be any combination of values
' in the range of the type supported. Subtracting 30 hours is
' equivalent to subtracting 1 day and 6 hours.

    Function = nSerial - TimeToSerial(uTime)
                       
End Function
' ========================================================================================
' Time Difference
' ========================================================================================
Private Sub cCalendar.SerialDifference (ByVal nFromSerial as LongInt, _
                                        ByVal nToSerial as LongInt, _
                                        ByRef uDiff as TIME_DURATION)
                     
Dim nTime   as Long

    cmSerialBreakApart(nToSerial - nFromSerial,uDiff.Days,nTime)
    cmTimeFromSerial(nTime,uDiff.Hour,uDiff.Minute,uDiff.Second,uDiff.Millisecond)
    uDiff.Days = Abs(uDiff.Days)                     
                     
End Sub
' ========================================================================================
' Time to Serial Time
' ========================================================================================
Private Function cCalendar.TimeToSerial (ByRef uTime as TIME_DURATION) as LongInt

Dim nTime   as Long

    nTime = cmTimeToSerial(uTime.Hour,uTime.Minute,uTime.Second,uTime.Millisecond)

    Function = (Abs(uTime.Days) * cCalendarClass.ONE_DAY + nTime) * IIf(uTime.Days < 0,-1,1)
    
End Function
' ========================================================================================
' Serial Time to Time
' ========================================================================================
Private Sub cCalendar.SerialToTime (ByVal nSerial as LongInt, _
                                    ByRef uTime as TIME_DURATION)
                 
Dim nTime   as Long

    cmSerialBreakApart(nSerial,uTime.Days,nTime)
    cmTimeFromSerial(nTime,uTime.Hour,uTime.Minute,uTime.Second,uTime.Millisecond)
   
End Sub

' ########################################################################################
' Unix Conversions
' ########################################################################################

' ========================================================================================
' Serial Time to Unix Time
' ========================================================================================
Private Function cCalendar.UnixFromSerial (ByVal nSerial as LongInt) as LongInt

' Unix keeps time as the number of seconds since Jan 1, 1970

' Ignore milliseconds

    Function = (nSerial - (cCalendarClass.UNIX_EPOCH * cCalendarClass.ONE_DAY)) _
             / cCalendarClass.ONE_SECOND

End Function
' ========================================================================================
' Unix Time to Serial Time
' ========================================================================================
Private Function cCalendar.SerialFromUnix (ByVal nUnix as LongInt) as LongInt

' Unix keeps time as the number of seconds since Jan 1, 1970

    Function = (nUnix + ((cCalendarClass.UNIX_EPOCH * cCalendarClass.ONE_DAY) _
             / cCalendarClass.ONE_SECOND)) * cCalendarClass.ONE_SECOND

End Function

' ########################################################################################
' Excel Conversions
' ########################################################################################

' ========================================================================================
' Serial Time to Excel Time
' ========================================================================================
Private Function cCalendar.ExcelFromSerial (ByVal nSerial as LongInt, _
                                            ByVal bBaseYear1904 as BOOLEAN) as Double

' Given a date and time, return an Excel date/time

' Excel dates are based on either Jan 1, 1900 or
' Jan 1, 1904. The Jan 1, 1904 base is used for
' compatibility with the Macintosh system.

' There is a known Excel date bug in that Feb 29, 1900
' is accepted as a valid date. This bug was probably inserted
' intentionally to be compatible with the same Lotus 123 bug.

' We will treat serial date 60 (2-29-1900) as equivalent to
' 3-1-1900 (serial date 61)

' bBaseYear1904 = TRUE or FALSE

Dim nMoment      as Double

    nMoment = cmSerialToMoment(nSerial)

    Select Case bBaseYear1904

    Case False

        nMoment = nMoment _
                - cCalendarClass.EXCEL_1900_EPOCH _
                + IIf(cmFloor(nMoment) > 60,1,0)_
                + 1

    Case Else

        nMoment = nMoment _
                - cCalendarClass.EXCEL_1904_EPOCH

    End Select

    Function = nMoment

End Function
' ========================================================================================
' Excel Time to Serial Time
' ========================================================================================
Private Function cCalendar.SerialFromExcel (ByVal nExcel as Double, _
                                            ByVal bBaseYear1904 as BOOLEAN) as LongInt

' Given an Excel serial date, return the date and time

' Excel dates are based on either Jan 1, 1900 or
' Jan 1, 1904. The Jan 1, 1904 base is used for
' compatibility with the Macintosh system.

' There is a known Excel date bug in that Feb 29, 1900
' is accepted as a valid date. This bug was probably inserted
' intentionally to be compatible with the same Lotus 123 bug.

' We will treat serial date 60 (2-29-1900) as equivalent to
' 2-28-1900 (serial date 59)

' bBaseYear1904 = TRUE or FALSE

Dim nMoment      as Double

    Select Case bBaseYear1904

    Case False

        nMoment = nExcel _
                + cCalendarClass.EXCEL_1900_EPOCH _
                - IIf(cmFloor(nExcel) < 60,0,1) _
                - 1

    Case Else

        nMoment = nExcel _
                + cCalendarClass.EXCEL_1904_EPOCH

    End Select

    Function = cmMomentToSerial(nMoment)

End Function

' ########################################################################################
' Rolling 13 Months Calculations
' ########################################################################################

' ========================================================================================
' Apply updates to history
' ========================================================================================
Private Function cCalendar.UpdateHistory (ByRef nHistoryMonth as Short, _
                                          ByRef nHistoryYear as Long, _
                                          arHistory() as HISTORY_MONTHS, _
                                          ByVal nUpdateMonth as Short, _
                                          ByVal nUpdateYear as Long, _
                                          arUpdates() as Double) as BOOLEAN

Dim nElaspedHistoryMonths     as Long
Dim nElaspedUpdateMonths      as Long
Dim nIndex                    as Long
Dim nMonth                    as Short
Dim nHistoryIndex             as Long

' If no history present, exit

    If UBound(arHistory) < 0 Then

        Function = False

        Exit Function

    End If

' History() and Updates() sizes must be the same

    If UBound(arHistory) <> UBound(arUpdates) Then

        Function = False

        Exit Function

    End If

' Convert Year/Month to elasped months

    nElaspedHistoryMonths = (nHistoryYear * 12) - 12 + nHistoryMonth
    nElaspedUpdateMonths = (nUpdateYear * 12) - 12 + nUpdateMonth

' Check if Update Year/Month are within last 12 months
' If not, no update is possible

    If nElaspedHistoryMonths - nElaspedUpdateMonths > 12 Then

        Function = True

        Exit Function

    End If

' Check if Update Year/Month are in the future and drop/clear months as needed

    Select Case nElaspedHistoryMonths - nElaspedUpdateMonths

    Case Is < 0

        For nHistoryIndex = 0 To UBound(arHistory)

            For nIndex = nElaspedHistoryMonths To nElaspedUpdateMonths - 1

                nMonth = cmFloor(cmAMod(nIndex,12))

                cmShiftHistory (nMonth,arHistory(nHistoryIndex))

            Next

        Next

        nHistoryMonth = nUpdateMonth

        nHistoryYear = nUpdateYear

        nElaspedHistoryMonths = nElaspedUpdateMonths

    End Select

' Calculate history month index

    nMonth = IIf(nElaspedHistoryMonths = nElaspedUpdateMonths,0,nUpdateMonth)

    For nHistoryIndex = 0 To UBound(arHistory)

        arHistory(nHistoryIndex).Month(nMonth) = arHistory(nHistoryIndex).Month(nMonth) _
                                               + arUpdates(nHistoryIndex)

    Next

    Function = True

End Function
' ========================================================================================
' Calculate the associated year(s) for the history months January - December
' ========================================================================================
Private Sub cCalendar.HistoryMonthsAndYears (ByVal nHistoryMonth as Long, _
                                             ByVal nHistoryYear as Long, _
                                             ByVal nFiscalYearBeginsMonth as Short, _
                                             ByVal bUseEndingMonth as BOOLEAN, _
                                             arFiscalYears() as Long)

Dim nMonth        as Short
Dim nYear         as Long

    ReDim arFiscalYears(11)

    For nMonth = 0 To 11

' Calculate the calendar year

        nYear = IIf(nMonth + 1 < nHistoryMonth,nHistoryYear,nHistoryYear - 1)

' Calculate the fiscal year

        arFiscalYears(nMonth) = GregorianFiscalYear(nMonth + 1,nYear,nFiscalYearBeginsMonth,bUseEndingMonth)

    Next

End Sub
' ========================================================================================
' Calculate Same Month Last Year History Summary
' ========================================================================================
Private Function cCalendar.SummarySameMonthLastYear (ByVal nHistoryMonth as Short, _
                                                     arHistory() as HISTORY_MONTHS, _
                                                     arSummary() as Double) as BOOLEAN

Dim nHistoryIndex     as Long

' If no history present, exit

    If UBound(arHistory) < 0 Then

        Function = False

        Exit Function

    End If

' Setup summary totals array

    ReDim arSummary(UBound(arHistory))

' Initialize summary totals

    cmClearSummary(arSummary())

' Add in month from last year

    For nHistoryIndex = 0 To UBound(arHistory)

        cmSumOneMonth(nHistoryMonth,arHistory(nHistoryIndex),arSummary(nHistoryIndex))

    Next

    Function = True

End Function
' ========================================================================================
' Calculate Last Quarter History Summary with support for fiscal years
' ========================================================================================
Private Function cCalendar.SummaryLastQuarter (ByVal nHistoryMonth as Short, _
                                               ByVal nStartFiscalMonth as Short, _
                                               arHistory() as HISTORY_MONTHS, _
                                               arSummary() as Double) as BOOLEAN

Dim nHistoryIndex     as Long
Dim nMonth            as Short
Dim nQuarterMonth     as Short

' If no history present, exit

    If UBound(arHistory) < 0 Then

        Function = False

        Exit Function

    End If

' Setup summary totals array

    ReDim arSummary(UBound(arHistory))

' Initialize summary totals

    cmClearSummary(arSummary())

' Last Quarter Totals are the three months prior
' to the first month of the current quarter.

' Calculate month (1-3) of the current quarter

    nQuarterMonth = cmFloor(cmAMod(GregorianFiscalMonth(nHistoryMonth, _
                                                        nStartFiscalMonth),3))

' Calculate last month of previous quarter

    nMonth = cmFloor(cmAMod(nHistoryMonth - nQuarterMonth,12))

' Add in the three months of the last fiscal quarter

    nQuarterMonth = 3

    While nQuarterMonth > 0

        For nHistoryIndex = 0 To UBound(arHistory)

            cmSumOneMonth(nMonth,arHistory(nHistoryIndex),arSummary(nHistoryIndex))

        Next

        nQuarterMonth = nQuarterMonth - 1

        nMonth = cmAMod(nMonth - 1,12)

    Wend

    Function = True

End Function
' ========================================================================================
' Calculate Quarter To Date History Summary with support for fiscal years
' ========================================================================================
Private Function cCalendar.SummaryQuarterToDate (ByVal nHistoryMonth as Short, _
                                                 ByVal nStartFiscalMonth as Long, _
                                                 arHistory() as HISTORY_MONTHS, _
                                                 arSummary() as Double) as BOOLEAN

Dim nHistoryIndex     as Long
Dim nMonth            as Short

' If no history present, exit

    If UBound(arHistory) < 0 Then

        Function = False

        Exit Function

    End If

' Setup summary totals array

    ReDim arSummary(UBound(arHistory))

' Initialize summary totals

    cmClearSummary(arSummary())

' QTD Totals include the current month (13) and
' all other previous months of the same or fiscal
' year quarter

' Calculate month (1-3) of the current quarter

    nMonth = cmAMod(GregorianFiscalMonth(nHistoryMonth, _
                                         nStartFiscalMonth),3)

' If history month is the first month of the fiscal quarter
' then no previous months of the same quarter are available

    Select Case nMonth

    Case Is <> 1

        nMonth = nMonth - 1

        While nMonth > 0

            For nHistoryIndex = 0 To UBound(arHistory)
 
                cmSumOneMonth(cmAMod(nHistoryMonth - nMonth,12),arHistory(nHistoryIndex),arSummary(nHistoryIndex))

            Next

            nMonth = nMonth - 1

        Wend

    End Select

' Add in current month

    For nHistoryIndex = 0 To UBound(arHistory)

        cmSumOneMonth(0,arHistory(nHistoryIndex),arSummary(nHistoryIndex))

    Next

    Function = True

End Function
' ========================================================================================
' Calculate Year To Date History Summary with support for fiscal years
' ========================================================================================
Private Function cCalendar.SummaryYearToDate (ByVal nHistoryMonth as Short, _
                                              ByVal nStartFiscalMonth as Short, _
                                              arHistory() as HISTORY_MONTHS, _
                                              arSummary() as Double) as Long

Dim nHistoryIndex     as Long
Dim nMonth            as Short

' If no history present, exit

    If UBound(arHistory) < 0 Then

        Function = False

        Exit Function

    End If

' Setup summary totals array

    ReDim arSummary(UBound(arHistory))

' Initialize summary totals

    cmClearSummary(arSummary())

' YTD Totals include the current month (0) and
' all other previous months of the same or fiscal
' year period

    Select Case nHistoryMonth

' IF history month is the first month of the fiscal year
' then no previous months of the same year are available

    Case Is = nStartFiscalMonth

    Case Else

' Add in previous months of the same fiscal year

        nMonth = nStartFiscalMonth

        While nMonth <> nHistoryMonth

            For nHistoryIndex = 0 To UBound(arHistory)

                cmSumOneMonth(nMonth,arHistory(nHistoryIndex),arSummary(nHistoryIndex))

            Next

            nMonth = IIf(nMonth = 12,1,nMonth + 1)

        Wend

    End Select

' Add in current month

    For nHistoryIndex = 0 To UBound(arHistory)

        cmSumOneMonth(0,arHistory(nHistoryIndex),arSummary(nHistoryIndex))

    Next

    Function = True

End Function

' ########################################################################################
' Date Calculation Support
' ########################################################################################

' ========================================================================================
' Apply any Fri/Sat/Sun/Year options to a date, results are returned in nObservedDays
' ========================================================================================
Private Function cCalendar.cmDateCalculationObserved (ByVal nThursdayOption as Short, _
                                                      ByVal nFridayOption as Short, _
                                                      ByVal nSaturdayOption as Short, _
                                                      ByVal nSundayOption as Short, _
                                                      ByVal nMondayOption as Short, _
                                                      ByVal nYearsOption as Short, _
                                                      ByVal nDays as Long, _
                                                      ByVal nYear as Long, _
                                                      ByRef nObservedDays as LongInt) as BOOLEAN

Dim nDayOfWeek        as Long
Dim nOddEven          as Long

    nObservedDays = nDays

    nDayOfWeek = cmGregorianWeekDay(nDays)
    
    Select Case nThursdayOption
    
    Case Is <> cCalendarClass.NO_THURSDAY_RULE
    
        Select Case nDayOfWeek
        
        Case cCalendarClass.THURSDAY

            If nThursdayOption = cCalendarClass.THURSDAY_OBSERVED_ON_WEDNESDAY Then
            
               nObservedDays = nObservedDays - 1
            
            End If 

        End Select

    End Select
    
    Select Case nFridayOption
    
    Case Is <> cCalendarClass.NO_FRIDAY_RULE
    
        Select Case nDayOfWeek
        
        Case cCalendarClass.FRIDAY
        
            If nFridayOption = cCalendarClass.FRIDAY_OBSERVED_ON_WEDNESDAY Then   
    
               nObservedDays = nObservedDays - 2
               
            Else
            
               If nFridayOption = cCalendarClass.FRIDAY_OBSERVED_ON_THURSDAY Then   
    
                  nObservedDays = nObservedDays - 1
                  
               End If
               
            End If
            
        End Select
    
    End Select

    Select Case nSaturdayOption

    Case Is <> cCalendarClass.NO_SATURDAY_RULE

' Check for Saturday

        Select Case nDayOfWeek

        Case cCalendarClass.SATURDAY

            If nSaturdayOption = cCalendarClass.SATURDAY_OBSERVED_ON_FRIDAY Then

                nObservedDays = nObservedDays - 1

            Else
            
               If nSaturdayOption = cCalendarClass.SATURDAY_OBSERVED_ON_SUNDAY Then

                   nObservedDays = nObservedDays + 1
                   
                Else

                   If nSaturdayOption = cCalendarClass.SATURDAY_OBSERVED_ON_MONDAY Then

                       nObservedDays = nObservedDays + 2

                   Else

                       If nSaturdayOption = cCalendarClass.SATURDAY_OBSERVED_ON_THURSDAY Then

                           nObservedDays = nObservedDays - 2
                           
                       End If

                    End If

                End If

            End If

        End Select

    End Select

    Select Case nSundayOption
    
    Case Is <> cCalendarClass.NO_SUNDAY_RULE
    
        Select Case nDayOfWeek
        
        Case cCalendarClass.SUNDAY

            If nSundayOption = cCalendarClass.SUNDAY_OBSERVED_ON_MONDAY Then
            
               nObservedDays = nObservedDays + 1
            
            End If 

        End Select

    End Select
    
    Select Case nMondayOption
    
    Case Is <> cCalendarClass.NO_MONDAY_RULE
    
        Select Case nDayOfWeek
        
        Case cCalendarClass.MONDAY

            If nMondayOption = cCalendarClass.MONDAY_OBSERVED_ON_TUESDAY Then
            
               nObservedDays = nObservedDays + 1
            
            End If 

        End Select

    End Select

' Check for year observation rules

    nOddEven = cmMod(Abs(nYear),2)

    Select Case nYearsOption

    Case cCalendarClass.ODD_YEARS_ONLY

        Function = IIf(nOddEven <> 0,True,False)

    Case cCalendarClass.EVEN_YEARS_ONLY

        Function = IIf(nOddEven = 0,True,False)
        
    Case cCalendarClass.YEARS_AFTER_LEAP_ONLY
    
        If cmJulianLeapYear(nYear - 1) = False Then
        
           Function = False
           
        Else
        
           Function = True 
        
        End If
    
    Case cCalendarClass.LEAP_YEARS_ONLY
    
        Function = cmJulianLeapYear(nYear)
        
    Case Else
    
        Function = True

    End Select

End Function
' ========================================================================================
' Save Date Calculation for business date calculations
' ========================================================================================
Private Sub cCalendar.cmSaveDateCalculation (ByRef uCalc as DATE_CALCULATION)

Dim nDays1          as Long
Dim nDays2          as Long
Dim bSaveDay1       as BOOLEAN
Dim bSaveDay2       as BOOLEAN
Dim bLoop           as BOOLEAN
Dim iIndex          as Long
Dim nObservedDays1  as LongInt
Dim nObservedDays2  as LongInt

    nObservedDays1 = uCalc.ObservedDays1
    nObservedDays2 = uCalc.ObservedDays2  

' Check for qualified dates

    Select Case uCalc.NonBusinessDate

       Case True
         
          If uCalc.Observed = True AndAlso nObservedDays1 <> nObservedDays2 Then

             nObservedDays1 = nObservedDays2

          End If

       Case Else

          Exit Sub

    End Select
     
    nDays1 = cmFloor(nObservedDays1 / cCalendarClass.ONE_DAY)
    nDays2 = cmFloor(nObservedDays2 / cCalendarClass.ONE_DAY)
     
' Check if dates are already saved

    bSaveDay1 = True
    bSaveDay2 = True 

    If UBound(This.arNonBusinessDays) >= 0 Then
    
       For iIndex = 0 To UBound(This.arNonBusinessDays)
       
           If This.arNonBusinessDays(iIndex) = nDays1 Then
              bSaveDay1 = False
           End If
           
           If This.arNonBusinessDays(iIndex) = nDays2 Then
              bSaveDay2 = False
           End If
           
' If neither date is to saved, exit loop

           If bSaveDay1 = False AndAlso bSaveDay1 = False Then
              Exit For
           End If
                 
       Next
    
    End If
    
' Save dates

    If bSaveDay1 = False AndAlso bSaveDay1 = False Then
    
       Exit Sub
       
    End If
    
' Check is limit on saved dates is reached

    If This.iUBound = UBound(This.arNonBusinessDays) Then
    
       uCalc.MaxNonBusinessDates = True
       Exit Sub    
    
    End If
    
    If bSaveDay1 = True Then
    
       This.iUBound = This.iUBound + 1
       This.arNonBusinessDays(iUBound) = nDays1
       
    End If
    
    If bSaveDay2 = True AndAlso bSaveDay1 <> bSaveDay2 Then
    
       This.iUBound = This.iUBound + 1
       This.arNonBusinessDays(iUBound) = nDays2
       
    End If
    
' Sort the list of dates to support bisection searches later
' A large list is not expected so a simple swap sort should be ok

' No sort necessary if only one date saved

    If This.iUBound > 0 Then
  
       bLoop = True
       
       While bLoop = True

        bloop = False
        
        For iIndex = 1 To This.iUBound
        
            If This.arNonBusinessDays(iIndex - 1) > This.arNonBusinessDays(iIndex) Then
            
               Swap This.arNonBusinessDays(iIndex - 1), This.arNonBusinessDays(iIndex)       
               bloop = True
               
            End If
            
        Next
    
       Wend
       
    End If   

End Sub

' ########################################################################################
' Easter Support
' ########################################################################################

' ========================================================================================
' Easter Calculations
' ========================================================================================
Private Sub cCalendar.cmEasterCalculation (ByVal nRuleClass as Short, _
                                           ByVal nRule as Short, _
                                           ByVal nYear as Long, _
                                           ByRef nCalcDays as Long)

Dim nDays  as Short

    Select Case nRuleClass

    Case cCalendarClass.CHRISTIANEASTER_RULES

        nCalcDays = cmChristianEasterDay(nYear)

        Select Case nRule

        Case cCalendarClass.CHRISTIAN_EASTER_GOODFRIDAY

            nDays = -2

        Case cCalendarClass.CHRISTIAN_EASTER_MAUNDYTHURSDAY

            nDays = -3

        Case cCalendarClass.CHRISTIAN_EASTER_PALMSUNDAY

            nDays = -7

        Case cCalendarClass.CHRISTIAN_EASTER_PASSIONSUNDAY

            nDays = -14

        Case cCalendarClass.CHRISTIAN_EASTER_MOTHERINGSUNDAY

            nDays = -21

        Case cCalendarClass.CHRISTIAN_EASTER_ASHWEDNESDAY

            nDays = -46

        Case cCalendarClass.CHRISTIAN_EASTER_MARDIGRAS

            nDays = -47

        Case cCalendarClass.CHRISTIAN_EASTER_SHROVEMONDAY

            nDays = -48

        Case cCalendarClass.CHRISTIAN_EASTER_SHROVESUNDAY

            nDays = -49

        Case cCalendarClass.CHRISTIAN_EASTER_SEXAGESIMASUNDAY

            nDays = -56

        Case cCalendarClass.CHRISTIAN_EASTER_EASTERMONDAY

            nDays = 1

        Case cCalendarClass.CHRISTIAN_EASTER_ROGATIONSUNDAY

            nDays = 35

        Case cCalendarClass.CHRISTIAN_EASTER_ASCENSION

            nDays = 39

        Case cCalendarClass.CHRISTIAN_EASTER_PENTECOST

            nDays = 49

        Case cCalendarClass.CHRISTIAN_EASTER_TRINITYSUNDAY

            nDays = 56

        Case cCalendarClass.CHRISTIAN_EASTER_CORPUSCHRISTI

            nDays = 60

        Case Else

            nDays = 0

        End Select

    Case Else

        nCalcDays = cmOrthodoxEasterDay(nYear)

        Select Case nRule

        Case cCalendarClass.ORTHODOX_EASTER_NEWYEAR

            nDays = 0
            nCalcDays = cmDaysFromJulian(cCalendarClass.JANUARY,1,nYear)

        Case cCalendarClass.ORTHODOX_EASTER_GOODFRIDAY

            nDays = -2

        Case cCalendarClass.ORTHODOX_EASTER_PALMSUNDAY

            nDays = -7

        Case cCalendarClass.ORTHODOX_EASTER_FORGIVENESSSUNDAY

            nDays = -49

        Case cCalendarClass.ORTHODOX_EASTER_GREATLENT

            nDays = -55

        Case cCalendarClass.ORTHODOX_EASTER_FEASTOFASCENSION

            nDays = 39

        Case cCalendarClass.ORTHODOX_EASTER_PENTECOST

            nDays = 49

        Case cCalendarClass.ORTHODOX_EASTER_APOSTLESFAST

            nDays = 50

        Case cCalendarClass.ORTHODOX_EASTER_ALLSAINTSSUNDAY

            nDays = 56

        Case Else

            nDays = 0

        End Select

    End Select

    nCalcDays = nCalcDays + nDays

End Sub
' ========================================================================================
' Calculate Easter for Orthodox churches
' ========================================================================================
Private Function cCalendar.cmOrthodoxEasterDay (ByVal nYear as Long) as Long

Dim nNicaeanRule           as Long
Dim nShiftedEpact          as Long

    nNicaeanRule = cmFloor(cmMod(nYear,19))

    nShiftedEpact = cmFloor(cmMod(14 + 11 * nNicaeanRule,30))

    Function = cmWeekDayAfter(cCalendarClass.SUNDAY,cmDaysFromJulian(cCalendarClass.APRIL,19,IIf(nYear > 0,nYear,nYear - 1)) - nShiftedEpact)

End Function
' ========================================================================================
' Calculate Easter for Catholic and Protestant churches
' ========================================================================================
Private Function cCalendar.cmChristianEasterDay (ByVal nYear as Long) as Long

Dim nPaschalMoon           as Long       ' Day after full moon on or after March 21
Dim nCenturies             as Long       ' Number of centuries in Easter year
Dim nNicaeanRule           as Long
Dim nLunarCycle            as Long
Dim nShiftedEpact          as Long       ' Age of moon (in days) for April 5

    nCenturies = cmFloor(nYear / 100) + 1

    nNicaeanRule = cmFloor(cmMod(nYear,19)) * 11 + 14      ' By Nicaean Rule (Orthodox Christian)

    nShiftedEpact = cmFloor(cmMod( _                       ' Correction for Metonic Cycle Inaccuracy
          nNicaeanRule - _
            cmFloor((3 * nCenturies) / 4) + _              ' Correction for Gregorian Centuries
            cmFloor((5 + 8 * nCenturies) / 25),30))        ' Correction for Metonic Cycle Inaccuracy

' Adjustment for 29.5 day lunar cycle

    nLunarCycle = cmFloor(cmMod(nYear,19))

    Select Case nShiftedEpact

    Case 0

        nShiftedEpact = nShiftedEpact + 1 

    Case 1

        nShiftedEpact = nShiftedEpact _
                      + IIf(10 < nLunarCycle,1,0)

    End Select

' nPaschalMoon = Day after full moon on or after March 21

    nPaschalMoon = cmDaysFromGregorian(cCalendarClass.APRIL,19,nYear) _
                 - nShiftedEpact

' Return the Sunday following nPaschalMoon

    Function =  cmWeekDayAfter(cCalendarClass.SUNDAY,nPaschalMoon)

End Function

' ########################################################################################
' ISO Support
' ########################################################################################

' ========================================================================================
' The ISO calendar has short (52) or long (53) week years.
' The last week of the year belongs to the year that has
' 4 or more days of the last week.
' ========================================================================================
Private Function cCalendar.cmISOLongYear (ByVal nYear as Long) as BOOLEAN

Dim nJan    as Long
Dim nDec    as Long

    cmGregorianYearRange(nYear,nJan,nDec)
   
    If cmGregorianWeekDay(nJan) = cCalendarClass.THURSDAY OrElse cmGregorianWeekDay(nDec) = cCalendarClass.THURSDAY Then
    
       Function = True
       
    Else
    
       Function = False
       
    End If  

End Function
' ========================================================================================
' Calculate ISO date from days
' ========================================================================================
Private Sub cCalendar.cmISOFromDays (ByVal nDays as Long, _
                                     ByRef nWeek as Short, _
                                     ByRef nDay as Short, _
                                     ByRef nYear as Long)

    nYear = cmGregorianYearFromDays(nDays - 3)

    Select Case nDays

        Case Is >= cmDaysFromISO(1,1,nYear + 1)

            nYear = nYear + 1

    End Select

    nWeek = cmFloor((nDays - cmDaysFromISO(1,1,nYear)) / 7) _
          + 1

    nDay = cmFloor(cmAMod(nDays,7))

End Sub
' ========================================================================================
' Calculate days date from ISO
' ========================================================================================
Private Function cCalendar.cmDaysFromISO (ByVal nWeek as Short, _
                                          ByVal nDay as Short, _
                                          ByVal nYear as Long) as Long

    Function = cmNthWeekDay(nWeek, _
                            cCalendarClass.SUNDAY, _
                            cmDaysFromGregorian(12,28,nYear - 1)) + nDay

End Function

' ########################################################################################
' Hebrew Support
' ########################################################################################

' ========================================================================================
' Calculate a Hebrew date based on rules
' ========================================================================================
Private Function cCalendar.cmHebrewDateCalculation (ByVal nMonth as Short, _
                                                    ByVal nDay as Short, _
                                                    ByVal nGregorianYear as Long, _
                                                    ByVal nWeekday as Short, _
                                                    ByVal nRule as Short, _
                                                    ByRef nCalcDays as Long) as BOOLEAN

' Return Hebrew holidays occurring in a given Gregorian year

Dim nJan1             as Long
Dim nDec31            as Long
Dim nHebrewYear       as Long

    nHebrewYear = nGregorianYear + 3760
    cmGregorianYearRange(nGregorianYear,nJan1,nDec31)

' Calculate the Hebrew Year

    nCalcDays = cmDaysFromHebrew(nMonth,nDay,nHebrewYear)

    Select Case nCalcDays

' Check if Hebrew date occurs during nGregorianYear

    Case nJan1 To nDec31

    Case Else

        nHebrewYear = nHebrewYear + 1

        nCalcDays = cmDaysFromHebrew(nMonth,nDay,nHebrewYear)

' Check if Hebrew date occurs at all during nGregorianYear

        Select Case nCalcDays

        Case nJan1 To nDec31

        Case Else

            Function = False

            Exit Function

        End Select

    End Select

' If month is Adar and we have a leap year, change month to AdarII

' Examples where this is applicable include Purim and Ta'anit Esther

    If nMonth = cCalendarClass.ADAR Then

        If cmHebrewLeapYear(nHebrewYear) = True Then

            nCalcDays = cmDaysFromHebrew(cCalendarClass.ADARII,nDay,nHebrewYear)

        End If

    End If

' Apply any rules

    Select Case nWeekday

    Case cCalendarClass.ALL_WEEKDAYS

        Select Case nRule

        Case cCalendarClass.HEBREW_HESHVAN30

' Examples using this rule: Rosh Chodesh Kislev

            If cmLastDayOfHebrewMonth(cCalendarClass.MARHESHVAN,nHebrewYear) = 30 Then

                nCalcDays = nCalcDays - 1

            End If

        Case cCalendarClass.HEBREW_KISLEV30

' Examples using this rule: Rosh Chodesh Tevet

            If cmLastDayOfHebrewMonth(cCalendarClass.KISLEV,nHebrewYear) = 30 Then

                nCalcDays = nCalcDays - 1

            End If

        End Select

    Case Else

' Check Rule

        Select Case nRule

            Case cCalendarClass.AFTER

' Examples using this rule: Shabbat Shuva - Saturday after New Year

                nCalcDays = cmWeekDayAfter(nWeekday,nCalcDays)

        End Select


    End Select

    Function = True

End Function
' ========================================================================================
' Hebrew Birthday in Hebrew Year
' ========================================================================================
Private Function cCalendar.cmHebrewBirthDay (ByVal nBirthMonth as Short, _
                                             ByVal nBirthDay as Short, _
                                             ByVal nBirthYear as Long, _
                                             ByVal nForHebrewYear as Long) as Long
                   
Dim nDays   as Long

    Select Case nBirthMonth
    
        Case Is = cmLastMonthOfHebrewYear(nBirthYear)
 
            nDays = cmDaysFromHebrew(cmLastMonthOfHebrewYear(nForHebrewYear), _
                                     nBirthDay, _
                                     nForHebrewYear)       
        
        Case Else

            nDays = cmDaysFromHebrew(nBirthMonth,1,nForHebrewYear) + nBirthDay - 1        
        
    End Select
                   
    Function = nDays               
                   
End Function
' ========================================================================================
' Calculate Hebrew Date from Days
' ========================================================================================
Private Sub cCalendar.cmHebrewFromDays (ByVal nDays as Long, _
                                        ByRef nMonth as Short, _
                                        ByRef nDay as Short, _
                                        ByRef nYear as Long)

' Year can be off by +-1

    nYear = cmFloor((98496 / 35975351) * (nDays - cCalendarClass.HEBREW_EPOCH))

    Do While cmHebrewNewYear(nYear) <= nDays

        nYear = nYear + 1

    Loop

    nYear = nYear - 1

' Starting Month for search

    nMonth = IIf(nDays < cmDaysFromHebrew(cCalendarClass.NISAN,1,nYear),cCalendarClass.TISHRI,cCalendarClass.NISAN)

' Look for Month that contains nDays

    Do While nDays > cmDaysFromHebrew(nMonth, _
                                      cmLastDayOfHebrewMonth(nMonth,nYear), _
                                      nYear)
        nMonth = nMonth + 1

    Loop

    nDay = nDays - cmDaysFromHebrew(nMonth,1,nYear) + 1

End Sub
' ========================================================================================
' Calculate Days from Hebrew Date
' ========================================================================================
Private Function cCalendar.cmDaysFromHebrew (ByVal nMonth as Short, _
                                             ByVal nDay as Short, _
                                             ByVal nYear as Long) as Long

Dim nDays             as Long
Dim nLastMonth        as Long
Dim nMonthsInYear     as Long

' Get start of the year plus days so far this month

    nDays = cmHebrewNewYear(nYear) _
          + nDay _
          - 1

' Add in elasped days in the given month and the length of each elasped month.
' Since Hebrew years begin on the seventh month (TISHRI), we have to check
' for months before and after TISHRI

    nMonthsInYear = cmLastMonthOfHebrewYear(nYear)

    Select Case nMonth

    Case Is < cCalendarClass.TISHRI

        For nLastMonth = cCalendarClass.TISHRI To nMonthsInYear

            nDays = nDays + cmLastDayOfHebrewMonth(nLastMonth,nYear)

        Next

        For nLastMonth = cCalendarClass.NISAN To nMonth - 1

            nDays = nDays + cmLastDayOfHebrewMonth(nLastMonth,nYear)

        Next

    Case Is > cCalendarClass.TISHRI

        For nLastMonth = cCalendarClass.TISHRI To nMonth - 1

            nDays = nDays + cmLastDayOfHebrewMonth(nLastMonth,nYear)

        Next

    End Select

    Function = nDays - IIf(nYear < 0,1,0)

End Function
' ========================================================================================
' Last Day of Hebrew Month
' ========================================================================================
Private Function cCalendar.cmLastDayOfHebrewMonth (ByVal nMonth as Long, _
                                                   ByVal nYear as Long) as Long

Dim nMonthLength       as Long

    nMonthLength = 30

' Look for 29 day months

    Select Case nMonth

    Case cCalendarClass.IYAR, cCalendarClass.TAMMUZ, cCalendarClass.ELUL, cCalendarClass.TEVET, cCalendarClass.ADARII

        nMonthLength = 29

    Case cCalendarClass.ADAR

        If cmHebrewLeapYear(nYear) = False Then

            nMonthLength = 29

        End If

    Case cCalendarClass.MARHESHVAN

        If cmLongMarheshvan(nYear) = False Then

            nMonthLength = 29

        End If

    Case cCalendarClass.KISLEV

        If cmShortKislev(nYear) = True Then

            nMonthLength = 29

        End If

     End Select

    Function = nMonthLength

End Function
' ========================================================================================
' Marheshvan Month Length
' ========================================================================================
Private Function cCalendar.cmLongMarheshvan (ByVal nYear as Long) as Long

    Select Case cmDaysInHebrewYear(nYear)

    Case 355, 385

        Function = True

    Case Else

        Function = False

    End Select

End Function
' ========================================================================================
' Kislev Month Length
' ========================================================================================
Private Function cCalendar.cmShortKislev (ByVal nYear as Long) as BOOLEAN

    Select Case cmDaysInHebrewYear(nYear)

    Case 353, 383

        Function = True

    Case Else

        Function = False

    End Select

End Function
' ========================================================================================
' Days in Hebrew Year
' ========================================================================================
Private Function cCalendar.cmDaysInHebrewYear (ByVal nYear as Long) as Long

    Function = cmHebrewNewYear(nYear + 1) _
             - cmHebrewNewYear(nYear)

End Function
' ========================================================================================
' Hebrew New Year
' ========================================================================================
Private Function cCalendar.cmHebrewNewYear (ByVal nYear as Long) as Long

    Function = cCalendarClass.HEBREW_EPOCH _
             + cmHebrewCalendarElapsedDays(nYear) _
             + cmHebrewYearLengthCorrection(nYear)

End Function
' ========================================================================================
' Add correction to length of Hebrew Year
' ========================================================================================
Private Function cCalendar.cmHebrewYearLengthCorrection (ByVal nYear as Long) as Long

Dim  nElaspedDays     as Long
Dim  nDays            as Long

    nElaspedDays = cmHebrewCalendarElapsedDays(nYear)

    If nElaspedDays - cmHebrewCalendarElapsedDays(nYear - 1) = 382 Then

        nDays = 1

    Else

        nDays = IIf(cmHebrewCalendarElapsedDays(nYear + 1) - nElaspedDays = 356,2,0)

    End If

    Function = nDays

End Function
' ========================================================================================
' Hebrew Elasped Days
' ========================================================================================
Private Function cCalendar.cmHebrewCalendarElapsedDays (ByVal nYear as Long) as Long

Dim nDays              as Long

    nDays = cmFloor(cmMolad(cCalendarClass.TISHRI,nYear) - cCalendarClass.HEBREW_EPOCH + .5)

    Function = IIf(cmMod(3 * (nDays + 1),7) < 3,nDays + 1,nDays)

End Function
' ========================================================================================
' Fixed moment of mean conjunction for Hebrew month and year
' ========================================================================================
Private Function cCalendar.cmMolad (ByVal nMonth as Short, _
                                    ByVal nYear as Long) as Double

Dim nMonthsElasped     as Long

    nMonthsElasped = nMonth _
                   - cCalendarClass.TISHRI _
                   + cmFloor((235 * IIf(nMonth < cCalendarClass.TISHRI,nYear + 1,nYear) - 234) / 19)

    Function = cCalendarClass.HEBREW_EPOCH _
             - (876 / 25920) _
             + (nMonthsElasped * (29.5 + (793 / 25920)))

End Function
' ========================================================================================
' Last month of Hebrew Year
' ========================================================================================
Private Function cCalendar.cmLastMonthOfHebrewYear (ByVal nYear as Long) as Short

    Function = IIf(cmHebrewLeapYear(nYear) = True,cCalendarClass.ADARII,cCalendarClass.ADAR)

End Function
' ========================================================================================
' Check for a Hebrew Leap Year
' ========================================================================================
Private Function cCalendar.cmHebrewLeapYear (ByVal nYear as Long) as BOOLEAN

    Function = IIf(cmMod((nYear * 7 + 1),19) < 7,True,False)

End Function
' ========================================================================================
' Check for a Hebrew Sabbatical Year
' ========================================================================================
Private Function cCalendar.cmHebrewSabbaticalYear (ByVal nYear as Long) as BOOLEAN

' Biblically mandated in Exodus 23:10-11

    Function = IIf(cmMod(nYear,7) = 0,True,False)
    
End Function

' ########################################################################################
' Persian Support
' ########################################################################################

' ========================================================================================
' Return Persian dates occurring in a given Gregorian year
' ========================================================================================
Private Function cCalendar.cmPersianDateCalculation (ByVal nMonth as Short, _
                                                     ByVal nDay as Short, _
                                                     ByVal nGregorianYear as Long, _
                                                     ByRef nCalcDays as Long) as BOOLEAN

Dim nJan1             as Long
Dim nDec31            as Long
Dim nPersianYear      as Long

    nPersianYear = cmPersianYear(nGregorianYear)
    cmGregorianYearRange(nGregorianYear,nJan1,nDec31)

' Calculate the Persian Date

    nCalcDays = cmDaysFromPersian(nMonth,nDay,nPersianYear)

    Select Case nCalcDays

' Check if Persian date occurs during nGregorianYear

    Case nJan1 To nDec31

        Function = True

    Case Else

        nPersianYear = nPersianYear - 1

        nCalcDays = cmDaysFromPersian(nMonth,nDay,nPersianYear)

' Check if Persian date occurs at all during nGregorianYear

        Select Case nCalcDays

        Case nJan1 To nDec31

            Function = True

        Case Else

            Function = False

        End Select

    End Select

End Function
' ========================================================================================
' Given a Gregorian year, return the Persian year equivalent
' ========================================================================================
Private Function cCalendar.cmPersianYear (ByVal nGregorianYear as Long) as Long

Dim nPersianYear      as Long

    nPersianYear = nGregorianYear - 621

' Compensate for the lack of year 0 on the Persian calendar
' We can still be off a year if Persian Month is >= 10

    Function = IIf(nPersianYear <= 0,nPersianYear - 1,nPersianYear)

End Function
' ========================================================================================
' Given a Days date, return Persian Date
' ========================================================================================
Private Sub cCalendar.cmPersianFromDays (ByVal nDays as Long, _
                       ByRef nMonth as Short, _
                       ByRef nDay as Short, _
                       ByRef nYear as Long)

Dim nNewYear       as Long
Dim nDayOfYear     as Long

    nNewYear = cmPersianNewYearOnOrBefore(nDays)
    nYear = cmRound((nNewYear - cCalendarClass.PERSIAN_EPOCH) / (cmMeanTropicalYear(0)) + 1)
    nYear = IIf(0 < nYear,nYear,nYear - 1)
    nDayOfYear = nDays - cmDaysFromPersian(cCalendarClass.FARVARDIN,1,nYear) + 1
    nMonth = IIf(nDayOfYear <= 186,cmCeiling(nDayOfYear / 31),cmCeiling((nDayOfYear - 6) / 30))
    nDay = nDays - cmDaysFromPersian(nMonth,1,nYear) + 1

End Sub
' ========================================================================================
' Given a Persian Date, return Days Date
' ========================================================================================
Private Function cCalendar.cmDaysFromPersian (ByVal nMonth as Short, _
                                              ByVal nDay as Short, _
                                              ByVal nYear as Long) as Long

Dim nDays             as Long

    nDays = cmPersianNewYearOnOrBefore(cCalendarClass.PERSIAN_EPOCH + 180 + cmFloor(cmMeanTropicalYear(0) * _
                                         IIf(0 < nYear,nYear - 1,nYear)))

    Function = nDays - 1 + IIf(nMonth <= 7,31 * (nMonth - 1),(30 * (nMonth - 1)) + 6) + nDay

End Function
' ========================================================================================
' Search for Persian New Year on vernal equinox (Around March 21)
' ========================================================================================
Private Function cCalendar.cmPersianNewYearOnOrBefore (ByVal nDays as Long) as Long

Dim nApprox        as Double

    nApprox = cmEstimatePriorSolarLongitude(cmMiddayInTehran(nDays),cCalendarClass.SPRING)

    If cmSolarLongitude(cmMiddayInTehran(cmFloor(nApprox))) > cCalendarClass.SPRING + 2 Then _

        nApprox = nApprox + 1

    End If

    Function = cmFloor(nApprox)

End Function
' ========================================================================================
' Midday or solar noon in Tehran
' ========================================================================================
Private Function cCalendar.cmMiddayInTehran (ByVal nDays as Long) as Double

    Function = cmUniversalFromStandard(cmMidday(nDays, _
                                                cCalendarClass.TehranLocale_Zone, _
                                                cCalendarClass.TehranLocale_Longitude), _
                                       cCalendarClass.TehranLocale_Zone)

End Function

' ########################################################################################
' Hindu Support
' ########################################################################################

' ========================================================================================
' Find Hindu Lunar Holiday in a given gregorian year
' ========================================================================================
Private Function cCalendar.cmHinduLunarDateCalculation (ByVal nRule as Short, _
                                                        ByVal nMonth as Short, _
                                                        ByVal nDay as Short, _
                                                        ByVal nGregorianYear as Long, _
                                                        ByRef nHoliday1 as Long, _
                                                        ByRef nHoliday2 as Long) as BOOLEAN

' The date may not occur or occur 1-2 times

Dim nJan1              as Long
Dim nDec31             as Long
Dim nMonth1            as Short
Dim bLeapMonth         as BOOLEAN
Dim nDay1              as Short
Dim bLeapDay           as BOOLEAN
Dim nYear1             as Long
Dim nMonth2            as Short
Dim nDay2              as Short
Dim nYear2             as Long
Dim bHolidayFound      as BOOLEAN

    nHoliday1 = 0
    nHoliday2 = 0
    bHolidayFound = False


' Calculate begin/end of gregorian year

    cmGregorianYearRange(nGregorianYear,nJan1,nDec31)

' Calculate hindu lunar begin/end of gregorian year

    cmHinduLunarFromDays(nJan1,nMonth1,bLeapMonth,nDay1,bLeapDay,nYear1)
    cmHinduLunarFromDays(nDec31,nMonth2,bLeapMonth,nDay2,bLeapDay,nYear2)
    nHoliday1 = cmAdjustedHindu(nMonth,False,nDay,False,nYear1)
    nHoliday2 = cmAdjustedHindu(nMonth,False,nDay,False,nYear2)

' Is estimated holiday date in current gregorian year
' and not a hindu skipped day or month?

    If nHoliday1 = 0 OrElse nHoliday1 < nJan1 OrElse nHoliday1 > nDec31 Then

       nHoliday1 = 0

    Else

       If cmExpunged(nMonth,nYear1) = False Then

          bHolidayFound = True

       Else

          nHoliday1 = 0

       End If
       
    End If
      
    If nHoliday2 = 0 OrElse nHoliday2 < nJan1 OrElse nHoliday2 > nDec31 Then

    Else

       If cmExpunged(nMonth,nYear2) = False Then

          If bHolidayFound = False Then

             nHoliday1 = nHoliday2
             bHolidayFound = True

          End If

      Else

          nHoliday2 = nHoliday1

      End If
      
    End If

    If bHolidayFound = True Then

        If nRule = cCalendarClass.BEFORE Then

' Example using this rule is Varlakshmi Vratam - Friday before Month 5, Day 15

            nHoliday1 = cmWeekDayBefore(cCalendarClass.FRIDAY,nHoliday1)
            nHoliday2 = cmWeekDayBefore(cCalendarClass.FRIDAY,nHoliday2)

        End If

    End If

    Function = bHolidayFound

End Function
' ========================================================================================
' Given a Hindu Lunar date, return the Days date
' ========================================================================================
Private Function cCalendar.cmDaysFromHinduLunar (ByVal nMonth as Short, _
                                                 ByVal bLeapMonth as BOOLEAN, _
                                                 ByVal nDay as Short, _
                                                 ByVal bLeapDay as BOOLEAN, _
                                                 ByVal nYear as Long) as Long

Dim nApprox                    as Double
Dim nEstimated                 as Long
Dim nSolarApprox               as Long
Dim nLunarDay                  as Long
Dim nAdjustment                as Long
Dim nMidMonth                  as Short
Dim bMidLeapMonth              as BOOLEAN
Dim nMidDay                    as Short
Dim bMidLeapDay                as BOOLEAN
Dim nMidYear                   as Long
Dim bLoop                      as BOOLEAN

' Rough Approximation

    nApprox = cCalendarClass.HINDU_EPOCH + cCalendarClass.HinduSiderealYear _
            * (nYear + cCalendarClass.HINDU_LUNAR_ERA + (nMonth - 1) / 12)

' Solar based approximation

    nSolarApprox = cmFloor(nApprox - (1 / 360) * cCalendarClass.HinduSiderealYear _
                 * (cmCalcDegrees(cmHinduSolarLongitude(nApprox) - (nMonth - 1) * 30 + 180)) - 180)

' Lunar Day of Solar Approximation

    nLunarDay = cmHinduLunarDayFromMoment(nSolarApprox + .25)

' Check for month

    If 3 < nLunarDay AndAlso nLunarDay < 27 Then          ' Borderline case - New Moon and Solar Month close to same

        nAdjustment = nLunarDay

    Else

' Middle of preceding solar month

        cmHinduLunarFromDays(nSolarApprox - 15,nMidMonth,bMidLeapMonth,nMidDay,bMidLeapDay,nMidYear)

' Look in preceding month

        If nMidMonth <> nMonth OrElse (bMidLeapMonth = True AndAlso bLeapMonth = False) Then

            nAdjustment = cmMod(nLunarDay + 15,30) - 15

        Else

' Look in next month

            nAdjustment = cmMod(nLunarDay - 15,30) + 15

        End If

    End If

' Calculate Estimation

    nEstimated = nSolarApprox + nDay - nAdjustment

' Refine Estimation

    nEstimated = nEstimated - cmMod(cmHinduLunarDayFromMoment(nEstimated + .25) - nDay + 15,30) + 15 - 1

    bLoop = True

    While bLoop = True

        cmHinduLunarFromDays(nEstimated,nMidMonth,bMidLeapMonth,nMidDay,bMidLeapDay,nMidYear)

        If cmHinduLunarOnOrBefore(nMonth,bLeapMonth,nDay,bLeapDay,nYear,nMidMonth,bMidLeapMonth,nMidDay,bMidLeapDay,nMidYear) = True Then

            bLoop = False

        Else

            nEstimated = nEstimated + 1

        End If

    Wend

    Function = nEstimated

End Function
' ========================================================================================
' Given a Days date, return the Hindu Lunar date
' ========================================================================================
Private Sub cCalendar.cmHinduLunarFromDays (ByVal nDays as Long, _
                                            ByRef nMonth as Short, _
                                            ByRef bLeapMonth as BOOLEAN, _
                                            ByRef nDay as Short, _
                                            ByRef bLeapDay as BOOLEAN, _
                                            ByRef nYear as Long)

Dim nCritical          as Double
Dim nLastNewMoon       as Double
Dim nNextNewMoon       as Double
Dim nSolarMonth        as Long

    nCritical = cmHinduSunRise(nDays)

' Set Day

    nDay = cmHinduLunarDayFromMoment(nCritical)

' Check for Leap Day

    bLeapDay = IIf(nDay = cmHinduLunarDayFromMoment(cmHinduSunRise(nDays - 1)),True,False)

' Calculate Last/Next New Moons and Solar Month

    nLastNewMoon = cmHinduNewMoonBefore(nCritical)
    nNextNewMoon = cmHinduNewMoonBefore(cmFloor(nLastNewMoon) + 35)
    nSolarMonth = cmHinduZodiac(nLastNewMoon)

' Set Month

    nMonth = cmAMod(nSolarMonth + 1,12)

' Check for Leap Month

    bLeapMonth = IIf(nSolarMonth = cmHinduZodiac(nNextNewMoon),True,False)

' Set Year
   
    nYear = cmHinduCalendarYear(IIf(nMonth <= 2,nNextNewMoon + 180,nNextNewMoon)) - cCalendarClass.HINDU_LUNAR_ERA

End Sub
' ========================================================================================
' Return Hindu Solar dates occurring in a given Gregorian year
' ========================================================================================
Private Function cCalendar.cmHinduSolarDateCalculation (ByVal nRule as Short, _
                                                        ByVal nMonth as Short, _
                                                        ByVal nDay as Short, _
                                                        ByVal nGregorianYear as Long, _
                                                        ByRef nCalcDays as Long) as BOOLEAN

Dim nJan1             as Long
Dim nDec31            as Long
Dim nHinduSolarYear   as Long

    Select Case nRule
    
    Case cCalendarClass.HINDU_SOLAR_MESHA_SAMKRANTI
    
       nCalcDays = cmFloor(cmHinduSolarLongitudeAtOrAfter(0,cmDaysFromGregorian(cCalendarClass.JANUARY,1,nGregorianYear)))
       
       Function = True
    
    Case Else

       nHinduSolarYear = nGregorianYear - 78
       cmGregorianYearRange(nGregorianYear,nJan1,nDec31)

' Calculate the Hindu Solar Date

       nCalcDays = cmDaysFromHinduSolar(nMonth,nDay,nHinduSolarYear)

       Select Case nCalcDays

' Check if HinduSolar date occurs during nGregorianYear

       Case nJan1 To nDec31

           Function = True

       Case Else

           nHinduSolarYear = nHinduSolarYear - 1

           nCalcDays = cmDaysFromHinduSolar(nMonth,nDay,nHinduSolarYear)

' Check if HinduSolar date occurs at all during nGregorianYear

           Select Case nCalcDays

           Case nJan1 To nDec31

               Function = True

           Case Else

               Function = False

           End Select
           
       End Select

    End Select

End Function
' ========================================================================================
' Given a Hindu Solar date, return a Days date
' ========================================================================================
Private Function cCalendar.cmDaysFromHinduSolar (ByVal nMonth as Short, _
                                                 ByVal nDay as Short, _
                                                 ByVal nYear as Long) as Long

Dim nBegin         as Long
Dim bLoop          as BOOLEAN

    nBegin = cmFloor((nYear + cCalendarClass.HINDU_SOLAR_ERA + ((nMonth - 1) / 12)) * cCalendarClass.HinduSiderealYear) _
            + cCalendarClass.HINDU_EPOCH - 3

    bLoop = True

    While bLoop = True

        If cmHinduZodiac(cmHinduSunRise(nBegin + 1)) <> nMonth Then

            nBegin = nBegin + 1

        Else

            bLoop = False

        End If

    Wend

    Function = nBegin + nDay - 1

End Function
' ========================================================================================
' Given a Days date, return the Hindu Solar date (Saka Era)
' ========================================================================================
Private Sub cCalendar.cmHinduSolarFromDays (ByVal nDays as Long, _
                                            ByRef nMonth as Short, _
                                            ByRef nDay as Short, _
                                            ByRef nYear as Long)

Dim nCritical      as Double
Dim nApprox        as Long
Dim bLoop          as BOOLEAN

    nCritical = cmHinduSunRise(nDays + 1)
    nMonth = cmHinduZodiac(nCritical)
    nYear = cmHinduCalendarYear(nCritical) - cCalendarClass.HINDU_SOLAR_ERA
    nApprox = nDays - 3 - cmMod(cmFloor(cmHinduSolarLongitude(nCritical)),30)
    bLoop = True

    While bLoop = True

        If cmHinduZodiac(cmHinduSunRise(nApprox + 1)) <> nMonth Then

            nApprox = nApprox + 1

        Else

            bLoop = False

        End If

    Wend

    nDay = nDays - nApprox + 1

End Sub
' ========================================================================================
' Check if Hindu Lunar Month is omitted
' ========================================================================================
Private Function cCalendar.cmExpunged (ByVal nMonth as Short, _
                                       ByVal nYear as Long) as BOOLEAN

Dim nDays               as Long
Dim nMonth1             as Short
Dim bLeapMonth1         as BOOLEAN
Dim nDay1               as Short
Dim bLeapDay1           as BOOLEAN
Dim nYear1              as Long

    nDays = cmDaysFromHinduLunar(nMonth,False,15,False,nYear)
    cmHinduLunarFromDays(nDays,nMonth1,bLeapMonth1,nDay1,bLeapDay1,nYear1)

    If nMonth <> nMonth1 Then

        Function = True

    Else

        Function = False

    End If

End Function
' ========================================================================================
' Adjust Days date for Hindu Lunar skipped months and days
' ========================================================================================
Private Function cCalendar.cmAdjustedHindu (ByVal nMonth as Short, _
                                            ByVal bLeapMonth as BOOLEAN, _
                                            ByVal nDay as Short, _
                                            ByVal bLeapDay as BOOLEAN, _
                                            ByVal nYear as Long) as Long

Dim nDays              as Long
Dim nMonth1            as Short
Dim bLeapMonth1        as BOOLEAN
Dim nDay1              as Short
Dim bLeapDay1          as BOOLEAN
Dim nYear1             as Long
Dim nMonth2            as Short
Dim bLeapMonth2        as BOOLEAN
Dim nDay2              as Short
Dim bLeapDay2          as BOOLEAN
Dim nYear2             as Long

    nDays = cmDaysFromHinduLunar(nMonth,bLeapMonth,nDay,bLeapDay,nYear)
    cmHinduLunarFromDays(nDays,nMonth1,bLeapMonth1,nDay1,bLeapDay1,nYear1)
    cmHinduLunarFromDays(nDays - 1,nMonth2,bLeapMonth2,nDay2,bLeapDay2,nYear2)

    If cmAlmostEqual(nMonth,bLeapMonth,nMonth1,bLeapMonth1) = True Then

        If bLeapDay = False OrElse bLeapDay1 = True Then

        Else

            nDays = nDays - 1

        End If

    Else

        If cmAlmostEqual(nMonth,bLeapMonth,nMonth2,bLeapMonth2) = True Then

            nDays = nDays - 1

        Else

' Bogus Date

            nDays = 0

        End If

    End If

    Function = nDays

End Function
' ========================================================================================
' Check if two Hindu Lunar dates are close to equal
' ========================================================================================
Private Function cCalendar.cmAlmostEqual (ByVal nMonth1 as Short, _
                                          ByVal bLeapMonth1 as BOOLEAN, _
                                          ByVal nMonth2 as Short, _
                                          ByVal bLeapMonth2 as BOOLEAN) as Long

    If bLeapMonth1 = bLeapMonth2 AndAlso nMonth1 = nMonth2 Then

        Function = True

    Else

        Function = False

    End If

End Function
' ========================================================================================
' Given two Hindu Lunar dates, determine if the first is on or before the second
' ========================================================================================
Private Function cCalendar.cmHinduLunarOnOrBefore (ByRef nMonth1 as Short, _
                                                   ByRef bLeapMonth1 as BOOLEAN, _
                                                   ByRef nDay1 as Short, _
                                                   ByRef bLeapDay1 as BOOLEAN, _
                                                   ByRef nYear1 as Long, _
                                                   ByRef nMonth2 as Short, _
                                                   ByRef bLeapMonth2 as BOOLEAN, _
                                                   ByRef nDay2 as Short, _
                                                   ByRef bLeapDay2 as BOOLEAN, _
                                                   ByRef nYear2 as Long) as BOOLEAN

Dim bReturn        as BOOLEAN

    bReturn = False

    Select Case nYear1

    Case Is > nYear2

    Case Else

        Select Case nMonth1

        Case Is > nMonth2

        Case Else

            If (bLeapMonth1 = True AndAlso bLeapMonth2 = False) OrElse _
                    (bLeapMonth1 = bLeapMonth2 AndAlso _
                        (nDay1 < nDay2 OrElse _
                            (nDay1 = nDay2 AndAlso _
                                (bLeapDay1 = False OrElse bLeapDay2 = True)))) Then

                bReturn = True

            End If

        End Select

    End Select

    Function = bReturn

End Function
' ========================================================================================
' Hindu Sunrise
' ========================================================================================
Private Function cCalendar.cmHinduSunRise (ByVal nDays as Long) as Double

' Uses modern computation of sunrise to get better agreement with published Hindu calendars
' instead of following the strict Hindu Surya-Siddhanta calculations which can be off by more
' than 16 minutes


Dim bBogus        as BOOLEAN

    Function = cmSunRise(nDays, _
                         cCalendarClass.HinduLocaleZone, _
                         cCalendarClass.HinduLocaleLatitude, _
                         cCalendarClass.HinduLocaleLongitude, _
                         cCalendarClass.HinduLocaleElevation, _
                         cCalendarClass.SUNRISE_SUNSET_TIME, _
                         bBogus)

End Function
' ========================================================================================
' Hindu Solar Calendar Year
' ========================================================================================
Private Function cCalendar.cmHinduCalendarYear (ByVal nMoment as Double) as Long

    Function = cmRound((((nMoment - cCalendarClass.HINDU_EPOCH) / cCalendarClass.HinduSiderealYear) _
             - (cmHinduSolarLongitude(nMoment) / 360)))

End Function
' ========================================================================================
' Hindu New Moon
' ========================================================================================
Private Function cCalendar.cmHinduNewMoonBefore (ByVal nMoment as Double) as Double

' Given a date/time moment, this routine will calculate the next
' new moon before nMoment.

' The basic strategy is to take the moment and search a bisection
' within an interval The search can terminate as soon as it has
' narrowed the position of the new moon down to one zodiacal sign
' and nExact = False. When nExact is True,the exact moment of the
' new moon is calculated.

Dim nStartMoment       as Double      ' Lower moment of range
Dim nEndMoment         as Double      ' Higher moment of range
Dim nNewMoment         as Double      ' bisection of Lower/Higher range
Dim nTolerance         as Double
Dim bLoop              as BOOLEAN

' Calculate bisection interval

    nNewMoment = nMoment _
                - (1 / 360) _
                * cmHinduLunarPhase(nMoment) _
                * cCalendarClass.HinduSynodicMonth
    nStartMoment = nNewMoment - 1
    nEndMoment = IIf(nMoment < nNewMoment + 1,nMoment,nNewMoment + 1)
    nNewMoment = (nEndMoment + nStartMoment) * .5
    nTolerance = 2^-1000
    bLoop = True

    While bLoop = True

          If cmHinduZodiac(nStartMoment) <> cmHinduZodiac(nEndMoment) OrElse _
            nEndMoment - nStartMoment < nTolerance Then

            If cmHinduLunarPhase(nNewMoment) < 180 Then

               nEndMoment = nNewMoment

            Else

               nStartMoment = nNewMoment

            End If

            nNewMoment = (nEndMoment + nStartMoment) * .5

        Else

            bLoop = False

        End If

    Wend

    Function = nNewMoment

End Function
' ========================================================================================
' Hindu Phase of the moon (tithi) at nMoment - returns values from 1 to 30
' ========================================================================================
Private Function cCalendar.cmHinduLunarDayFromMoment (ByVal nMoment as Double) as Long

    Function = cmFloor(cmHinduLunarPhase(nMoment) / 12) + 1

End Function
' ========================================================================================
' Hindu Lunar Phase
' ========================================================================================
Private Function cCalendar.cmHinduLunarPhase (ByVal nMoment as Double) as Double

    Function = cmCalcDegrees(cmHinduLunarLongitude(nMoment) - cmHinduSolarLongitude(nMoment))

End Function
' ========================================================================================
' Hindu Lunar Longitude
' ========================================================================================
Private Function cCalendar.cmHinduLunarLongitude (ByVal nMoment as Double) as Double

    Function = cmHinduTruePosition(nMoment,cCalendarClass.HinduSiderealMonth,32 / 360, _
                                   cCalendarClass.HinduAnomalisticMonth,1 / 96)

End Function
' ========================================================================================
' Hindu Zodiac Sign
' ========================================================================================
Private Function cCalendar.cmHinduZodiac (ByVal nMoment as Double) as Long

    Function = cmFloor(cmHinduSolarLongitude(nMoment) / 30) + 1

End Function
' ========================================================================================
' Time at or after nMoment when solar longitude will be target
' ========================================================================================
Private Function cCalendar.cmHinduSolarLongitudeAtOrAfter (ByVal nTargetLongitude as Double, _
                                         ByVal nMoment as Double) as Double

Dim nTau                as Double
Dim nStartMoment        as Double
Dim nEndMoment          as Double
Dim nNewMoment          as Double      ' bisection of Lower/Higher range
Dim nNewLongitude       as Double
Dim bLoop               as BOOLEAN

    nTau = nMoment + cCalendarClass.HinduSiderealYear * ( 1 / 360) _
         * (cmCalcDegrees((nTargetLongitude - cmHinduSolarLongitude(nMoment))))

' Estimate within 5 days
         
    nStartMoment = IIf(nMoment > nTau - 5,nMoment,nTau - 5)
    nEndMoment = nTau + 5
    bLoop = True
    
    While bLoop = True

        nNewMoment = nStartMoment + ((nEndMoment - nStartMoment) * .5)
        nNewLongitude = cmCalcDegrees(cmHinduSolarLongitude(nNewMoment) - nTargetLongitude)

        Select Case nNewLongitude

        Case Is < 180

            nEndMoment = nNewMoment

        Case Else

            nStartMoment = nNewMoment

        End Select

        bLoop = IIf(nEndMoment - nStartMoment >= .000001,True,False)

    Wend

    Function = nStartMoment + ((nEndMoment - nStartMoment) * .5)    
                                         
End Function
' ========================================================================================
' Hindu Solar Longitude
' ========================================================================================
Private Function cCalendar.cmHinduSolarLongitude (ByVal nMoment as Double) as Double

    Function = cmHinduTruePosition(nMoment,cCalendarClass.HinduSiderealYear,14 / 360, _
                                   cCalendarClass.HinduAnomalisticYear,1 / 42)

End Function
' ========================================================================================
' Longitudinal Solar or Lunar position at nMoment
' ========================================================================================
Private Function cCalendar.cmHinduTruePosition (ByVal nMoment as Double, _
                                                ByVal nPeriod as Double, _                ' Period of mean motions
                                                ByVal nSize as Double, _                  ' Ratio of radii of epicycle and deferent
                                                ByVal nAnomalistic as Double, _           ' Period of retrograde revolution about epicycle
                                                ByVal nChange as Double) as Double        ' Maximum decrease in epicycle size

Dim nLong          as Double       ' Position of epicycle center
Dim nOffset        as Double       ' Sine of Anomaly
Dim nContraction   as Double
Dim nEquation      as Double       ' Equation of center

    nLong = cmHinduMeanPosition(nMoment,nPeriod)
    nOffset = cmHinduSine(cmHinduMeanPosition(nMoment,nAnomalistic))
    nContraction = Abs(nOffset) * nChange * nSize
    nEquation = cmHinduArcsin(nOffset * (nSize - nContraction))

    Function = cmCalcDegrees(nLong - nEquation)

End Function
' ========================================================================================
' Position in degrees at nMoment in uniform circular orbit of nPeriod days
' ========================================================================================
Private Function cCalendar.cmHinduMeanPosition (ByVal nMoment as Double, _
                                                ByVal nPeriod as Double) as Double

    Function = 360 * Frac((nMoment - cCalendarClass.HinduCreation) / nPeriod)

End Function
' ========================================================================================
' Inverse of cmHinduSine()
' ========================================================================================
Private Function cCalendar.cmHinduArcsin (ByVal nAmp as Double) as Double

Dim nPosition      as Long
Dim bNegative      as BOOLEAN
Dim bLoop          as BOOLEAN
Dim bReturn        as Double
Dim nBelow         as Double

    nPosition = 0
    bNegative = IIf(nAmp < 0,True,False)
    nAmp = IIf(bNegative = True,nAmp * -1,nAmp)
    bLoop = True

    While bLoop = True

        If nAmp > cmHinduSineTable(nPosition) Then

            nPosition = nPosition + 1

        Else

            bLoop = False

        End If

    Wend

    nBelow = cmHinduSineTable(nPosition - 1)
    bReturn = (225 / 60) * (nPosition - 1 + ((nAmp - nBelow) / (cmHinduSineTable(nPosition) - nBelow)))

    Function = IIf(bNegative = True,bReturn * -1,bReturn)

End Function
' ========================================================================================
' Linear interpolation of Hindu Sine Table
' ========================================================================================
Private Function cCalendar.cmHinduSine (ByVal nDegrees as Double) as Double

Dim nFraction      as Double
Dim nEntry         as Double

    nEntry = nDegrees / (225 / 60)
    nFraction = Frac(nEntry)

    Function = nFraction * cmHinduSineTable(cmCeiling(nEntry)) _
             + (1 - nFraction) * cmHinduSineTable(cmFloor(nEntry))
End Function
' ========================================================================================
' Hindu Sine Table simulation where nEntry is a multiplier of 225 minutes
' ========================================================================================
Private Function cCalendar.cmHinduSineTable (ByVal nEntry as Long) as Double

Dim nExact         as Double
Dim nError         as Double

    nExact = 3438 * cmSinDegrees(nEntry * (225 / 60))
    nError = 0.215 * cmSignum(nExact) * cmSignum(Abs(nExact) - 1716)

    Function = cmRound(nExact + nError) / 3438.0

End Function

' ########################################################################################
' Islamic Support
' ########################################################################################

' ========================================================================================
' Return Islamic dates occurring in a given Gregorian year
' ========================================================================================
Private Function cCalendar.cmIslamicDateCalculation (ByVal nRule as Short, _
                                                     ByVal nMonth as Short, _
                                                     ByVal nDay as Short, _
                                                     ByVal nGregorianYear as Long, _
                                                     ByRef nHoliday1 as Long, _
                                                     ByRef nHoliday2 as Long) as BOOLEAN

Dim nDays               as Long
Dim nWorkDays1          as Long
Dim nWorkDays2          as Long
Dim bFirstDate          as BOOLEAN
Dim bFirstValidDate     as BOOLEAN
Dim bSecondValidDate    as BOOLEAN

    nHoliday1 = 0
    nHoliday2 = 0

    Select Case nRule

    Case cCalendarClass.ISLAMIC_QUDS_DAY

' Quds Day - Last Friday in Ramadan (Month 9)

' Get days date of first Day of Month 10

        cmIslamicInGregorian(cCalendarClass.SHAWWAL,1,nGregorianYear,nWorkDays1,bFirstValidDate,nWorkDays2,bSecondValidDate)

        If bFirstValidDate = True Then

           nDays = cmWeekDayBefore(cCalendarClass.FRIDAY,nWorkDays1)

           If cmGregorianYearFromDays(nDays) = nGregorianYear Then

              nHoliday1 = nDays
              nHoliday2 = nDays
              
           Else
           
              bFirstValidDate = False

           End If
           
        End If
           
        If bSecondValidDate = True Then

           nDays = cmWeekDayBefore(cCalendarClass.FRIDAY,nWorkDays2)

           If cmGregorianYearFromDays(nDays) = nGregorianYear Then

              If bFirstValidDate = True Then

                 nHoliday2 = nDays

              Else

                 bFirstValidDate = True
                 nHoliday1 = nDays
                 nHoliday2 = nDays
                    
              End If

           End If

        End If

    Case Else

        cmIslamicInGregorian(nMonth,nDay,nGregorianYear,nHoliday1,bFirstValidDate,nHoliday2,bSecondValidDate)

    End Select

    Function = bFirstValidDate

End Function
' ========================================================================================
' Given a Days Date, return Islamic Date
' ========================================================================================
Private Function cCalendar.cmDaysFromIslamic (ByVal nMonth as Short, _
                                              ByVal nDay as Short, _
                                              ByVal nYear as Long) as Long

Dim nMidMonth      as Long

    nMidMonth = cCalendarClass.ISLAMIC_EPOCH _
              + cmFloor(((nYear - 1) * 12 + nMonth - .5) * cCalendarClass.MeanSynodicMonth)

    Function = cmPhasisOnOrBefore(nMidMonth) _
             + nDay - 1

End Function
' ========================================================================================
' Given an Islamic Date, return Days Date
' ========================================================================================
Private Sub cCalendar.cmIslamicFromDays (ByVal nDays as Long, _
                                         ByRef nMonth as Short, _
                                         ByRef nDay as Short, _
                                         ByRef nYear as Long)

Dim nCrescent          as Long
Dim nElaspedMonths     as Long

    nCrescent = cmPhasisOnOrBefore(nDays)

    nElaspedMonths = cmRound((nCrescent - cCalendarClass.ISLAMIC_EPOCH) / cCalendarClass.MeanSynodicMonth)
    nYear = cmFloor(nElaspedMonths / 12) + 1
    nMonth = cmMod(nElaspedMonths,12) + 1
    nDay = nDays - nCrescent + 1

End Sub
' ========================================================================================
' Determine days dates of Islamic date in a Gregorian year
' ========================================================================================
Private Sub cCalendar.cmIslamicInGregorian (ByVal nIslamicMonth as Short, _
                                            ByVal nIslamicDay as Short, _
                                            ByVal nGregorianYear as Long, _
                                            ByRef nFirstDate as Long, _
                                            ByRef bFirstValidDate as BOOLEAN, _
                                            ByRef nSecondDate as Long, _
                                            ByRef bSecondValidDate as BOOLEAN)

' The date always occurs at least once and possibly twice.

' A Gregorian year contains parts of at least two Islamic
' years and possibly three.

Dim nGregorianStart        as Long
Dim nGregorianEnd          as Long
Dim nIslamicWorkMonth      as Short
Dim nIslamicWorkDay        as Short
Dim nIslamicWorkYear       as Long
Dim nDays                  as Long

    bFirstValidDate = False
    bSecondValidDate = False
    nFirstDate = 0
    nSecondDate = 0
    
    cmGregorianYearRange(nGregorianYear,nGregorianStart,nGregorianEnd)

' Check first possible Islamic Year

    cmIslamicFromDays(nGregorianStart,nIslamicWorkMonth,nIslamicWorkDay,nIslamicWorkYear)
    nDays = cmDaysFromIslamic(nIslamicMonth,nIslamicDay,nIslamicWorkYear)

    If nDays >= nGregorianStart AndAlso nDays <= nGregorianEnd Then

       bFirstValidDate = True
       nFirstDate = nDays
       nSecondDate = nDays

    End If

' Check second possible Islamic Year

    nDays = cmDaysFromIslamic(nIslamicMonth,nIslamicDay,nIslamicWorkYear + 1)

    If nDays >= nGregorianStart AndAlso nDays <= nGregorianEnd Then

        If bFirstValidDate = False Then

           bFirstValidDate = True
           nFirstDate = nDays
           nSecondDate = nDays

        Else

           bSecondValidDate = True
           nSecondDate = nDays

        End If

    End If

' Check third possible Islamic Year

    If bFirstValidDate = True AndAlso bSecondValidDate = True Then

    Else

        nDays = cmDaysFromIslamic(nIslamicMonth,nIslamicDay,nIslamicWorkYear + 2)

        If nDays >= nGregorianStart AndAlso nDays <= nGregorianEnd Then

            If bFirstValidDate = False Then

               nFirstDate = nDays
               nSecondDate = nDays
               bFirstValidDate = True

            Else

               nSecondDate = nDays

            End If

        End If

    End If

End Sub
' ========================================================================================
' Crescent moon
' ========================================================================================
Private Function cCalendar.cmPhasisOnOrBefore (ByVal nDays as Long) as Long

' Uses a universal format where, if the UTC time of a new moon if <= Noon
' a crescent moon must be visible in the world somewhere on or after the
' the international date line else it will be visible the next day

Dim nMoment       as Double
Dim nCrescent     as Long

    nMoment = cmLunarPhaseAtOrBefore(nDays,cCalendarClass.NEWMOON) + 1
    nCrescent = cmFloor(nMoment) + IIf(Frac(Abs(nMoment)) > .5,1,0)

    If nCrescent > nDays Then

        nMoment = cmLunarPhaseAtOrBefore(nCrescent - 15,cCalendarClass.NEWMOON) + 1
        nCrescent = cmFloor(nMoment) + IIf(Frac(Abs(nMoment)) > .5,1,0)

    End If

    Function = nCrescent

End Function

' ########################################################################################
' Coptic Support
' ########################################################################################

' ========================================================================================
' Return Coptic dates occurring in a given Gregorian year
' ========================================================================================
Private Function cCalendar.cmCopticDateCalculation (ByVal nMonth as Short, _
                                                    ByVal nDay as Short, _
                                                    ByVal nGregorianYear as Long, _
                                                    ByRef nCalcDays as Long) as BOOLEAN

Dim nJan1             as Long
Dim nDec31            as Long
Dim nCopticMonth      as Short
Dim nCopticDay        as Short
Dim nCopticYear       as Long

    cmGregorianYearRange(nGregorianYear,nJan1,nDec31)
    cmCopticFromDays(nJan1,nCopticMonth,nCopticDay,nCopticYear)

' Calculate the Coptic Date

    nCalcDays = cmDaysFromCoptic(nMonth,nDay,nCopticYear)

    Select Case nCalcDays

' Check if Coptic date occurs during nGregorianYear

    Case nJan1 To nDec31

    Function = True

    Case Else

        nCopticYear = nCopticYear + 1

        nCalcDays = cmDaysFromCoptic(nMonth,nDay,nCopticYear)

' Check if Coptic date occurs at all during nGregorianYear

        Select Case nCalcDays

        Case nJan1 To nDec31

            Function = True

        Case Else

            Function = False

        End Select

    End Select

End Function
' ========================================================================================
' Given a Days Date, return Coptic Date
' ========================================================================================
Private Sub cCalendar.cmCopticFromDays (ByVal nDays as Long, _
                                        ByRef nMonth as Short, _
                                        ByRef nDay as Short, _
                                        ByRef nYear as Long)

    nYear = cmFloor((4 * (nDays - cCalendarClass.COPTIC_EPOCH) + 1463) / 1461)

    nMonth = cmFloor((nDays - cmDaysFromCoptic(1,1,nYear)) / 30) + 1

    nDay = nDays + 1 - cmDaysFromCoptic(nMonth,1,nYear)

End Sub
' ========================================================================================
' Given a Coptic Date, return Days Date
' ========================================================================================
Private Function cCalendar.cmDaysFromCoptic(ByVal nMonth as Short, _
                                            ByVal nDay as Short, _
                                            ByVal nYear as Short) as Long

    Function = cmFloor(cCalendarClass.COPTIC_EPOCH - 1 + 365 * (nYear - 1) _
             + (cmFloor(nYear / 4)) _
             + 30 * (nMonth - 1) + nDay)

End Function

' ########################################################################################
' Ethiopic Support
' ########################################################################################

' ========================================================================================
' Return Ethiopic dates occurring in a given Gregorian year
' ========================================================================================
Private Function cCalendar.cmEthiopicDateCalculation (ByVal nMonth as Short, _
                                                      ByVal nDay as Short, _
                                                      ByVal nGregorianYear as Long, _
                                                      ByRef nCalcDays as Long) as BOOLEAN

Dim nJan1             as Long
Dim nDec31            as Long
Dim nEthiopicMonth    as Short
Dim nEthiopicDay      as Short
Dim nEthiopicYear     as Long

    cmGregorianYearRange(nGregorianYear,nJan1,nDec31)
    cmEthiopicFromDays(nJan1,nEthiopicMonth,nEthiopicDay,nEthiopicYear)

' Calculate the Ethiopic Date

    nCalcDays = cmDaysFromEthiopic(nMonth,nDay,nEthiopicYear)

    Select Case nCalcDays

' Check if Ethiopic date occurs during nGregorianYear

    Case nJan1 To nDec31

        Function = True

    Case Else

        nEthiopicYear = nEthiopicYear + 1

        nCalcDays = cmDaysFromEthiopic(nMonth,nDay,nEthiopicYear)

' Check if Ethiopic date occurs at all during nGregorianYear

        Select Case nCalcDays

        Case nJan1 To nDec31

            Function = True

        Case Else

            Function = False

        End Select

    End Select

End Function
' ========================================================================================
' Given a Days Date, return Ethioptic Date
' ========================================================================================
Private Sub cCalendar.cmEthiopicFromDays (ByVal nDays as Long, _
                                          ByRef nMonth as Short, _
                                          ByRef nDay as Short, _
                                          ByRef nYear as Long)

    cmCopticFromDays(nDays + cCalendarClass.COPTIC_EPOCH - cCalendarClass.ETHIOPIC_EPOCH, _
                     nMonth, _
                     nDay, _
                     nYear)

End Sub
' ========================================================================================
' Given a Ethioptic Date, return Days Date
' ========================================================================================
Private Function cCalendar.cmDaysFromEthiopic(ByVal nMonth as Short, _
                                              ByVal nDay as Short, _
                                              ByVal nYear as Short) as Long

    Function = cCalendarClass.ETHIOPIC_EPOCH _
             + cmDaysFromCoptic(nMonth,nDay,nYear) _
             - cCalendarClass.COPTIC_EPOCH

End Function

' ########################################################################################
' Roman Support
' ########################################################################################

' ========================================================================================
' Calculate Roman Date from Days
' ========================================================================================
Private Sub cCalendar.cmRomanFromDays (ByVal nDays as Long, _
                                       ByRef nMonth as Short, _
                                       ByRef nYear as Long, _
                                       ByRef nEvent as Short, _
                                       ByRef nCount as Short, _
                                       ByRef bLeap as BOOLEAN)
                    
' First step is to convert to a Julian date and then deciding how to proceed

Dim nJulianMonth    as Short
Dim nJulianDay      as Short
Dim nJulianYear     as Long

    cmJulianFromDays(nDays,nJulianMonth,nJulianDay,nJulianYear)
    
    Select Case nJulianDay
    
        Case 1
        
            nMonth = nJulianMonth
            nYear = nJulianYear
            nEvent = cCalendarClass.KALENDS
            nCount = 1
            bLeap = False
        
        Case Is <= cmNonesOfMonth(nJulianMonth)
        
            nMonth = nJulianMonth
            nYear = nJulianYear
            nEvent = cCalendarClass.NONES
            nCount = cmNonesOfMonth(nJulianMonth) - nJulianDay + 1
            bLeap = False
        
        Case Is <= cmIdesOfMonth(nJulianMonth)
        
            nMonth = nJulianMonth
            nYear = nJulianYear
            nEvent = cCalendarClass.IDES
            nCount = cmIdesOfMonth(nJulianMonth) - nJulianDay + 1
            bLeap = False
        
        Case Else
        
            If nJulianMonth <> cCalendarClass.FEBRUARY OrElse cmJulianLeapYear(nJulianYear) = False Then

                nMonth = cmAMod(nJulianMonth + 1,12)
                If nMonth <> 1 Then
                
                    nYear = nJulianYear 
                
                Else
                
                    If nJulianYear <> -1 Then
                    
                        nYear = nJulianYear + 1
                    
                    Else
                    
                        nYear = 1
                        
                    End If
                    
                End If
                
                nEvent = cCalendarClass.KALENDS
                nCount = cmDaysFromRoman(nMonth,nYear,cCalendarClass.KALENDS,1,False) - nDays + 1
                bLeap = False
                
            Else
            
                If nJulianDay < 25 Then
                
                   nMonth = cCalendarClass.MARCH
                   nYear = nJulianYear
                   nEvent = cCalendarClass.KALENDS
                   nCount = 30 - nJulianDay
                   bLeap = False 
                
                Else                

                   nMonth = cCalendarClass.MARCH
                   nYear = nJulianYear
                   nEvent = cCalendarClass.KALENDS
                   nCount = 31 - nJulianDay
                   bLeap = IIf(nJulianDay = 25,True,False) 
                 
                End If    
                
            End If
        
    End Select    
                    
End Sub
' ========================================================================================
' Calculate Days from Roman Date
' ========================================================================================
Private Function cCalendar.cmDaysFromRoman (ByVal nMonth as Short, _
                                            ByVal nYear as Long, _
                                            ByVal nEvent as Short, _
                                            ByVal nCount as Short, _
                                            ByVal bLeap as BOOLEAN) as Long

Dim nDays   as Long
Dim nAdjust as Short

    Select Case nEvent
    
        Case cCalendarClass.KALENDS

            nDays = cmDaysFromJulian(nMonth,1,nYear)
            
        Case cCalendarClass.NONES
        
            nDays = cmDaysFromJulian(nMonth,cmNonesOfMonth(nMonth),nYear)
            
        Case Else
        
            nDays = cmDaysFromJulian(nMonth,cmIdesOfMonth(nMonth),nYear)
            
    End Select
    
    If cmJulianLeapYear(nYear) = True AndAlso _
            nMonth = cCalendarClass.MARCH AndAlso _
            nEvent = cCalendarClass.KALENDS AndAlso _
            16 >= nCount AndAlso _
            nCount >= 6 Then
            
       nAdjust = 0
       
    Else
    
       nAdjust = 1
       
    End If

    Function = nDays + nAdjust + IIf(bLeap = True,1,0) - nCount                     
                         
End Function
' ========================================================================================
' Roman Nones for month
' ========================================================================================
Private Function cCalendar.cmNonesOfMonth (ByVal nMonth as Short) as Short

    Function = cmIdesOfMonth(nMonth) - 8
    
End Function
' ========================================================================================
' Roman Ides for month
' ========================================================================================
Private Function cCalendar.cmIdesOfMonth (ByVal nMonth as Short) as Short

    Select Case nMonth
    
        Case cCalendarClass.MARCH,cCalendarClass.MAY,cCalendarClass.JULY,cCalendarClass.OCTOBER
        
            Function = 15
            
        Case Else
        
            Function = 13
            
    End Select

End Function

' ########################################################################################
' Armenian Support
' ########################################################################################

' ========================================================================================
' Calculate Armenian Date from Days
' ========================================================================================
Private Function cCalendar.cmDaysFromArmenian (ByVal nMonth as Short, _
                                               ByVal nDay as Short, _
                                               ByVal nYear as Long) as Long

' Calculate days date from Armenian date

    Function = cCalendarClass.ARMENIAN_EPOCH + cmDaysFromEgyptian(nMonth,nDay,nYear) _
             - cCalendarClass.EGYPTIAN_EPOCH

End Function
' ========================================================================================
' Calculate Days from Armenian Date
' ========================================================================================
Private Sub cCalendar.cmArmenianFromDays (ByVal nDays as Long, _
                                          ByRef nMonth as Short, _
                                          ByRef nDay as Short, _
                                          ByRef nYear as Long)

' Calculate Armenian date from days

    cmEgyptianFromDays(nDays + cCalendarClass.EGYPTIAN_EPOCH - _
                       cCalendarClass.ARMENIAN_EPOCH,nMonth,nDay,nYear)

End Sub

' ########################################################################################
' Egyptian Support
' ########################################################################################

' ========================================================================================
' Calculate Egyptian Date from Days
' ========================================================================================
Private Function cCalendar.cmDaysFromEgyptian (ByVal nMonth as Short, _
                                               ByVal nDay as Short, _
                                               ByVal nYear as Long) as Long

' Calculate days date from Egyptian date

    Function = cCalendarClass.EGYPTIAN_EPOCH + 365 * (nYear - 1) + 30 * (nMonth - 1) + nDay - 1

End Function
' ========================================================================================
' Calculate Days from Egyptian Date
' ========================================================================================
Private Sub cCalendar.cmEgyptianFromDays (ByVal nDays as Long, _
                                          ByRef nMonth as Short, _
                                          ByRef nDay as Short, _
                                          ByRef nYear as Long)

' Calculate Egyptian date from days

Dim nCalcDays   as Long

    nCalcDays = nDays - cCalendarClass.EGYPTIAN_EPOCH
    nYear = cmFloor(nCalcDays / 365) + 1
    nMonth = cmFloor((1 / 30) * (cmMod(nCalcDays,365))) + 1
    nDay = nCalcDays - 365 * (nYear - 1) - 30 * (nMonth - 1) + 1

End Sub

' ########################################################################################
' Bahai Support
' ########################################################################################

' ========================================================================================
' Return Bahai dates occurring in a given Gregorian year
' ========================================================================================
Private Function cCalendar.cmBahaiDateCalculation (ByVal nRule as Short, _
                                                   ByVal nMonth as Short, _
                                                   ByVal nDay as Short, _
                                                   ByVal nGregorianYear as Long, _
                                                   ByRef nCalcDays as Long) as BOOLEAN

Dim nJan1             as Long
Dim nDec31            as Long
Dim nNewYear          as Double
Dim nNewMoonAfter     as Double
Dim nNewMoon          as Double
Dim nSunset           as Double
Dim iMoon             as Short
Dim uBahai            as BAHAI_DATE
Dim nYears            as Long

    nNewYear = cmBahaiNewYearOnOrBefore(cmDaysFromGregorian(cCalendarClass.MARCH,28,nGregorianYear))
    cmGregorianYearRange(nGregorianYear,nJan1,nDec31)

    Select Case nRule

       Case cCalendarClass.NAW_RUZ

          nCalcDays = cmFloor(nNewYear)

       Case Is = cCalendarClass.BIRTH_OF_BAB, cCalendarClass.BIRTH_OF_BAHAULLAH 

' Definition is the day following the 8th new moon after the new year

          nNewMoonAfter = cmNewMoonAfter(nNewYear - 1) 
          nNewMoon = nNewYear

          For iMoon = 1 To 8 - IIf(nNewMoonAfter > nNewYear,0,1)
    
             nNewMoon = cmNewMoonAfter(nNewMoon + 1)
       
          Next

' Since a Bahai day begins at sunset on the day before a Gregorian date, check if the
' new moon was after sunset.

          nSunset = cmBahaiSunset(cmFloor(nNewMoon))

          nCalcDays = cmFloor(nNewMoon + IIf(nNewMoon > nSunset,1,0) + 1) _
                    + IIf(nRule = cCalendarClass.BIRTH_OF_BAHAULLAH,1,0)

       Case Else
       
          uBahai.Month = nMonth
          uBahai.Day = nDay
          nYears = nGregorianYear - cmGregorianYearFromDays(cCalendarClass.BAHAI_EPOCH)
          uBahai.Major = cmFloor(nYears / 361) + 1
          uBahai.Cycle = cmFloor((1 / 19) * cmMod(nYears,361)) + 1
          uBahai.Year = cmMod(nYears,19) + 1 
          nCalcDays = cmDaysFromBahai(uBahai.Major,uBahai.Cycle,uBahai.Year,uBahai.Month,uBahai.Day)
          
' Check if the date is in the requested Gregorian Year

          Select Case nCalcDays
          
             Case nJan1 To nDec31
                
             Case Else

                nYears = nYears - 1
                uBahai.Major = cmFloor(nYears / 361) + 1
                uBahai.Cycle = cmFloor((1 / 19) * cmMod(nYears,361)) + 1
                uBahai.Year = cmMod(nYears,19) + 1 
                nCalcDays = cmDaysFromBahai(uBahai.Major,uBahai.Cycle,uBahai.Year,uBahai.Month,uBahai.Day)
                
          End Select             

    End Select    

    Function = True

End Function
' ========================================================================================
' Calculate Bahai Date from Days
' ========================================================================================
Private Sub cCalendar.cmBahaiFromDays (ByVal nDays as Long, _
                                       ByRef nMajor as Short, _
                                       ByRef nCycle as Short, _
                                       ByRef nMonth as Short, _
                                       ByRef nDay as Short, _
                                       ByRef nYear as Short)

Dim nNewYear        as Long
Dim nYears          as Long
Dim nYearDays       as Long

    nNewYear = cmFloor(cmBahaiNewYearOnOrBefore(nDays))
    nYears = cmRound((nNewYear - cCalendarClass.BAHAI_EPOCH) / cmMeanTropicalYear(0))

    nMajor = cmFloor(nYears / 361) + 1
    nCycle = cmFloor((1 / 19) * (cmMod(nYears,361))) + 1
    nYear = cmMod(nYears,19) + 1
    nYearDays = nDays - nNewYear
    
    If nDays >= cmDaysFromBahai(nMajor,nCycle,nYear,19,1) Then
    
       nMonth = 19
       
    Else
    
       If nDays >= cmDaysFromBahai(nMajor,nCycle,nYear,cCalendarClass.AYYAMIHA,1) Then
       
          nMonth = cCalendarClass.AYYAMIHA
          
       Else
       
          nMonth = cmFloor(nYearDays / 19) + 1
          
       End If
       
    End If      

    nDay = nDays + 1 _
         - cmDaysFromBahai(nMajor,nCycle,nYear,nMonth,1)

End Sub
' ========================================================================================
' Calculate Days from Bahai Date
' ========================================================================================
Private Function cCalendar.cmDaysFromBahai (ByVal nMajor as Short, _
                                            ByVal nCycle as Short, _
                                            ByVal nYear as Short, _
                                            ByVal nMonth as Short, _
                                            ByVal nDay as Short) as Long

Dim nYears      as Double
Dim nDays       as Long

    nYears = 361 * (nMajor - 1) + 19 * (nCycle - 1) + nYear
    
    Select Case nMonth
    
        Case cCalendarClass.ALA
 
            nDays = cmFloor(cmBahaiNewYearOnOrBefore(cCalendarClass.BAHAI_EPOCH + cmFloor(cmMeanTropicalYear(0)) _
                  * (nYears + .5))) _
                  + nDay _
                  - 20       
        
        Case cCalendarClass.AYYAMIHA

            nDays = cmFloor(cmBahaiNewYearOnOrBefore(cCalendarClass.BAHAI_EPOCH + cmFloor(cmMeanTropicalYear(0)) _
                  * (nYears - .5))) _
                  + nDay _
                  + 341       
        
        Case Else

            nDays = cmFloor(cmBahaiNewYearOnOrBefore(cCalendarClass.BAHAI_EPOCH + cmFloor(cmMeanTropicalYear(0)) _
                  * (nYears - .5))) _
                  + ((nMonth - 1) * 19) _
                  + nDay _
                  - 1        
    End Select        

    Function = nDays                               
                         
End Function
' ========================================================================================
' Search for Bahai New Year on vernal equinox (Around March 21)
' ========================================================================================
Private Function cCalendar.cmBahaiNewYearOnOrBefore (ByVal nDays as Long) as Double

' The first day of Bahai Badi calendar is the day on which the vernal
' equinox occurs before sunset in Tehran

Dim nSunset        as Double
Dim nSpring        as Double  
Dim nMonth         as Short
Dim nDay           as Short
Dim nYear          as Long
Dim nSpringYear    as Long    

    cmGregorianFromDays(nDays,nMonth,nDay,nYear)
    nSpringYear = nYear
    nSpring = cmSeasonalEquinox(nYear,cCalendarClass.SPRING)
   
    If cmFloor(nSpring) > nDays Then

        nSpringYear = nYear - 1
        nSpring = cmSeasonalEquinox(nSpringYear,cCalendarClass.SPRING)     
    
    End If
    
    nSunset = cmBahaiSunset(cmFloor(nSpring))
    
' Check for vernal equinox after sunset when nDays = nSpring is found
    
    If cmFloor(nSpring) = nDays AndAlso nSpring > nSunset Then

       nSpring = cmSeasonalEquinox(nSpringYear - 1,cCalendarClass.SPRING)     
       nSunset = cmBahaiSunset(cmFloor(nSpring))
    
    End If

    Function = nSpring + IIf(nSpring > nSunset,1,0)

End Function
' ========================================================================================
' Bahai Sunset in Tehran
' ========================================================================================
Private Function cCalendar.cmBahaiSunset (ByVal nDays as Long) as Double

' Since the solar longitude calculations are more accurate than
' the sunset calculations, we make a small adjustment (about 8.5 minutes) to bring
' them closer together for the years when sunset in Tehran and the vernal equinox
' are very close together such as in Gregorian year 2026.

Dim bBogus         as BOOLEAN

    Function = cmUniversalFromStandard( _
                 cmSunSet(nDays, _
                          cCalendarClass.TehranLocale_Zone, _
                          cCalendarClass.TehranLocale_Latitude, _
                          cCalendarClass.TehranLocale_Longitude, _
                          cCalendarClass.TehranLocale_Elevation, _
                          cCalendarClass.SUNRISE_SUNSET_TIME, _
                          bBogus), _
                 cCalendarClass.TehranLocale_Zone)  - .006

End Function

' ########################################################################################
' Tibetan Support
' ########################################################################################

' ========================================================================================
' Return Tibetan dates occurring in a given Gregorian year
' ========================================================================================
Private Function cCalendar.cmTibetanDateCalculation (ByVal nRule as Short, _
                                                     ByVal nMonth as Short, _
                                                     ByVal nDay as Short, _
                                                     ByVal nGregorianYear as Long, _
                                                     ByRef nCalcDays as Long) as BOOLEAN

Dim nJan1             as Long
Dim nDec31            as Long

    Select Case nRule

       Case cCalendarClass.LOSAR

          nCalcDays = cmTibetanNewYearInGregorian(nGregorianYear)

       Case Else

          cmGregorianYearRange(nGregorianYear,nJan1,nDec31)
          nCalcDays = cmDaysFromTibetan(nMonth,False,nDay,False,nGregorianYear + 127)
          
' Check if the date is in the requested Gregorian Year

          Select Case nCalcDays
          
             Case nJan1 To nDec31
                
             Case Else

                cmDaysFromTibetan(nMonth,False,nDay,False,nGregorianYear + 126)
                
          End Select             

    End Select    

    Function = True

End Function
' ========================================================================================
' For a given Gregorian year, return Tibetan New Year
' ========================================================================================
Private Function cCalendar.cmTibetanNewYearInGregorian (ByVal nGregorianYear as Long) as Long

Dim uTibetan      as TIBETAN_DATE
Dim nDec31        as Long
Dim nDays         as Long

    nDec31 = cmGregorianYearEnd(nGregorianYear)
    cmTibetanFromDays(nDec31,uTibetan.Month,uTibetan.LeapMonth,uTibetan.Day,uTibetan.LeapDay,uTibetan.Year)
    nDays = cmTibetanLosar(uTibetan.Year)

    Function = nDays

End Function
' ========================================================================================
' For a given Tibetan year, return Losar
' ========================================================================================
Private Function cCalendar.cmTibetanLosar (ByVal nYear as Long) as Long

    Function = cmDaysFromTibetan(cCalendarClass.DBO, _
                                 cmTibetanLeapMonth(cCalendarClass.DBO,nYear), _
                                 1, _
                                 False, _
                                 nYear)

End Function
' ========================================================================================
' Check if a Tibetan Month in a Tibetan Year is a leap month
' ========================================================================================
Private Function cCalendar.cmTibetanLeapMonth (ByVal nMonth as Short, _
                                               ByVal nYear as Long) as BOOLEAN
                                               
Dim uTibetan    as TIBETAN_DATE

    cmTibetanFromDays(cmDaysFromTibetan(nMonth,True,2,False,nYear), _
                      uTibetan.Month,uTibetan.LeapMonth,uTibetan.Day,uTibetan.LeapDay,uTibetan.Year)
    
    Function = IIf(nMonth = uTibetan.Month,True,False)                                              

End Function
' ========================================================================================
' Given a Days date, return the Tibetan date
' ========================================================================================
Private Sub cCalendar.cmTibetanFromDays (ByVal nDays as Long, _
                                         ByRef nMonth as Short, _
                                         ByRef bLeapMonth as BOOLEAN, _
                                         ByRef nDay as Short, _
                                         ByRef bLeapDay as BOOLEAN, _
                                         ByRef nYear as Long)

Dim nTibetanMeanYear   as Double
Dim nYears             as Long
Dim nYear0             as Long
Dim nMonth0            as Short
Dim nDay0              as Short
Dim bLoop              as BOOLEAN
Dim nEstimated         as Long

    nTibetanMeanYear = 365 + 4975 / 18382
    nYears = cmCeiling((nDays - cCalendarClass.TIBETAN_EPOCH) / nTibetanMeanYear)
    bLoop = True
    nYear0 = nYears
    
    While bLoop = True
    
        If nDays >= cmDaysFromTibetan(cCalendarClass.DBO,False,1,False,nYear0) Then
        
           nYear0 = nYear0 + 1
           
        Else
        
           bLoop = False
           
        End If
    
    Wend

    nYear0 = nYear0 - 1
    
    nMonth0 = cCalendarClass.DBO
    bLoop = True
    
    While bLoop = True

        If nDays >= cmDaysFromTibetan(nMonth0,False,1,False,nYear0) Then
        
           nMonth0 = nMonth0 + 1
           
        Else
        
           bLoop = False
           
        End If    
    
    Wend

    nMonth0 = nMonth0 - 1

    nEstimated = nDays - cmDaysFromTibetan(nMonth0,False,1,False,nYear0)
    nDay0 = nEstimated - 2
    bLoop = True
    
    While bLoop = True

        If nDays >= cmDaysFromTibetan(nMonth0,False,nDay0,False,nYear0) Then
        
          nDay0 = nDay0 + 1
           
        Else
        
          bLoop = False
           
        End If    
    
    Wend

    nDay0 = nDay0 - 1

' Determine if leap month

    If nDay0 > 30 Then
    
       bLeapMonth = True
       
    Else
    
       bLeapMonth = False
       
    End If

' Determine day of month
    
    nDay = cmAMod(nDay0,30)

' Determine month of year

    If nDay > nDay0 Then

       nMonth = nMonth0 - 1

    Else

       If bLeapMonth = True Then

          nMonth = nMonth0 + 1

       Else

          nMonth = nMonth0

       EndIf

    End If

    nMonth = cmAMod(nMonth,12)

' Determine year

    If nDay > nDay0 AndAlso nMonth0 = 1 Then

       nYear = nYear0 - 1

    Else

       If bLeapMonth = True AndAlso nMonth0 = 12 Then

          nYear = nYear0 + 1

       Else

          nYear = nYear0

       End If

    End If

' Determine leap day

    If nDays = cmDaysFromTibetan(nMonth,bLeapMonth,nDay,True,nYear) Then

       bLeapDay = True

    Else

       bLeapDay = False

    End If

End Sub
' ========================================================================================
' Given a Tibetan date, return the Days date
' ========================================================================================
Private Function cCalendar.cmDaysFromTibetan (ByVal nMonth as Short, _
                                              ByVal bLeapMonth as BOOLEAN, _
                                              ByVal nDay as Short, _
                                              ByVal bLeapDay as BOOLEAN, _
                                              ByVal nYear as Long) as Long
Dim nMonths         as Double
Dim nDays           as Double
Dim nMean           as Double
Dim nSolarAnomaly   as Double
Dim nLunarAnomaly   as Double
Dim nSun            as Double
Dim nMoon           as Double 

    nMonths = cmFloor(((804 / 65) * (nYear - 1)) + ((67 / 65) * nMonth) + (IIf(bLeapMonth = True,-1,0) + (64 / 65)))

' Lunar Day count    
    
    nDays = (nMonths * 30) + nDay
    
' Mean civil days since epoch    
    
    nMean = (nDays * (11135 / 11312)) - 30 + (1071 / 1616) + IIf(bLeapDay = True,0,-1)
    
    nSolarAnomaly = cmMod((nDays * (13 / 4824)) + (2117 / 4824),1)
    nLunarAnomaly = cmMod((nDays * (3781 / 105840)) + (2837 / 15120),1)
    nSun = cmTibetanSunEquation(nSolarAnomaly * 12)
    nMoon = cmTibetanMoonEquation(nLunarAnomaly * 28)

    Function = cmFloor(cCalendarClass.TIBETAN_EPOCH + nMean + nSun + nMoon)

End Function
' ========================================================================================
' Tibetan interpolated tabular sine of solar anomaly
' ========================================================================================
Private Function cCalendar.cmTibetanSunEquation (ByVal nAlpha as Double) as Double
     
    Select Case nAlpha

       Case Is > 6 

          Function = cmTibetanSunEquation(nAlpha - 6) * -1

       Case Is > 3

          Function = cmTibetanSunEquation(6 - nAlpha)

       Case Else

          Select Case nAlpha
          
             Case 0
             
                Function = 0
                
             Case 1
             
                Function = 6 / 60
                
             Case 2
             
                Function = 10 / 60
                
             Case 3
             
                Function = 11 / 60
                
             Case Else

                Function = (cmMod(nAlpha,1) * cmTibetanSunEquation(cmCeiling(nAlpha))) _
                         + (cmMod(nAlpha * -1,1) * cmTibetanSunEquation(cmFloor(nAlpha)))             
             
          End Select   
        
    End Select    

End Function
' ========================================================================================
' Tibetan interpolated tabular sine of lunar anomaly
' ========================================================================================
Private Function cCalendar.cmTibetanMoonEquation (ByVal nAlpha as Double) as Double

    Select Case nAlpha

       Case Is > 14

          Function = cmTibetanMoonEquation(nAlpha - 14) * -1

       Case Is > 7

          Function = cmTibetanMoonEquation(14 - nAlpha)
        
       Case Else 
          
          Select Case nAlpha
          
             Case 0
             
                Function = 0
                
             Case 1
             
                Function = 5 / 60
                
             Case 2
             
                Function = 10 / 60
                
             Case 3
             
                Function = 15 / 60
                
             Case 4
             
                Function = 19 / 60
                
             Case 5
             
                Function = 22 / 60
                
             Case 6
                   
                Function = 24 / 60
                
             Case 7
             
                Function = 25 / 60
                
             Case Else
             
                Function = (cmMod(nAlpha,1) * cmTibetanMoonEquation(cmCeiling(nAlpha))) _
                         + (cmMod(nAlpha * -1,1) * cmTibetanMoonEquation(cmFloor(nAlpha)))             
             
          End Select          
                   
    End Select    

End Function

' ########################################################################################
' Gregorian Support
' ########################################################################################

' ========================================================================================
' Calculate a Gregorian date based on rules
' ========================================================================================
Private Sub cCalendar.cmGregorianDateCalculation (ByVal nMonth as Short, _
                                                  ByVal nDay as Short, _
                                                  ByVal nYear as Long, _
                                                  ByVal nWeekDay as Short, _
                                                  ByVal nRule as Short, _
                                                  ByRef nCalcDays as Long)

' Output is: nCalcDays,nCalcMonth, nCalcDay, nCalcYear

' nYear: -9999 through 9999
' nMonth: 1 through 12
' nDay: 1-31, varies based on
'                nCalcMonth - only used when
'                        nWeekday = -1

' nWeekday: 0=Sun...6=Sat, -1 = Any weekday

' If nCalcWeekday is Sun through Sat then:

'   nRule: If nWeekday is Sun through Sat then:

'           1=First,2=Second,3=Third,4=Fourth,5=Last,6=Last Full
'           7=Before,8=OnOrBefore,9=After,10=OnOrAfter,11=Nearest

'   nRule: If nWeekday is -1 then:

'           1=Before,2=OnOrBefore,3=After,4=OnOrAfter,5=Nearest,6=No Rule

Dim nDayOfWeek        as Short

' Check Weekday option

    Select Case nWeekDay

    Case cCalendarClass.ALL_WEEKDAYS

        nCalcDays = cmDaysFromGregorian(nMonth,nDay,nYear)

    Case Else

' Check Rule

        Select Case nRule

        Case cCalendarClass.FIRST_WEEK To cCalendarClass.FOURTH_WEEK

            nCalcDays = cmDaysFromGregorian(nMonth,1,nYear)

        Case cCalendarClass.LAST_WEEK

            nCalcDays = cmDaysFromGregorian(nMonth, _
                                            GregorianDaysInMonth(nMonth,nYear), _
                                            nYear)

        Case cCalendarClass.LAST_FULL_WEEK

' Check if last day of month falls on a Saturday
' If not, back up to Saturday of last full week
' and use on or before logic to find date

            nCalcDays = cmDaysFromGregorian(nMonth, _
                                            GregorianDaysInMonth(nMonth,nYear), _
                                            nYear)
            nDayOfWeek = cmGregorianWeekDay(cmDaysFromGregorian(nMonth, _
                                                                GregorianDaysInMonth(nMonth,nYear), _
                                                                nYear))
            nRule = cCalendarClass.ON_OR_BEFORE
            nCalcDays = IIf(nDayOfWeek = cCalendarClass.SATURDAY,nCalcDays,nCalcDays - nDayOfWeek - 1)

        Case Else

' We must be calculating a date before/after/nearest to some base date

            nCalcDays = cmDaysFromGregorian(nMonth,nDay,nYear)

        End Select

    End Select

' nCalcDays now contains our base date for further calculations

    Select Case nRule

    Case cCalendarClass.FIRST_WEEK

        nCalcDays = cmFirstWeekDay(nWeekDay,nCalcDays)

    Case cCalendarClass.SECOND_WEEK

       nCalcDays = cmSecondWeekDay(nWeekDay,nCalcDays)

    Case cCalendarClass.THIRD_WEEK

       nCalcDays = cmThirdWeekDay(nWeekDay,nCalcDays)

    Case cCalendarClass.FOURTH_WEEK

       nCalcDays = cmFourthWeekDay(nWeekDay,nCalcDays)

    Case cCalendarClass.LAST_WEEK

       nCalcDays = cmLastWeekDay(nWeekDay,nCalcDays)

    Case cCalendarClass.BEFORE

       nCalcDays = cmWeekDayBefore(nWeekDay,nCalcDays)

    Case cCalendarClass.ON_OR_BEFORE

      nCalcDays = cmWeekDayOnOrBefore(nWeekDay,nCalcDays)

    Case cCalendarClass.AFTER

       nCalcDays = cmWeekDayAfter(nWeekDay,nCalcDays)

    Case cCalendarClass.ON_OR_AFTER

       nCalcDays = cmWeekDayOnOrAfter(nWeekDay,nCalcDays)

    Case cCalendarClass.NEAREST

       nCalcDays = cmWeekDayNearest(nWeekDay,nCalcDays)

    End Select

End Sub
' ========================================================================================
' Calculate the difference in days between two Gregorian dates
' ========================================================================================
Private Function cCalendar.cmGregorianDateDifference (ByVal nStartMonth as Short, _
                                                      ByVal nStartDay as Short, _
                                                      ByVal nStartYear as Long, _
                                                      ByVal nEndMonth as Short, _
                                                      ByVal nEndDay as Short, _
                                                      ByVal nEndYear as Long) as Long
                                 
    Function = cmDaysFromGregorian(nEndMonth,nEndDay,nEndYear) _
             - cmDaysFromGregorian(nStartMonth,nStartDay,nStartYear)

End Function
' ========================================================================================
' Return the range of days in a gregorian year in days format
' Used to determine if holidays on other calendars fall
' during a particular Gregorian Year
' ========================================================================================
Private Sub cCalendar.cmGregorianYearRange (ByVal nYear as Long, _
                                            ByRef nYearStart as Long, _
                                            ByRef nYearEnd as Long)

    nYearStart = cmGregorianNewYear(nYear)
    nYearEnd = cmGregorianYearEnd(nYear)

End Sub
' ========================================================================================
' Return the first day of a gregorian year in days format
' ========================================================================================
Private Function cCalendar.cmGregorianNewYear (ByVal nYear as Long) as Long

    Function = cmDaysFromGregorian(cCalendarClass.JANUARY, _
                                   1, _
                                   nYear)

End Function
' ========================================================================================
' Return the last day of a gregorian year in days format
' ========================================================================================
Private Function cCalendar.cmGregorianYearEnd (ByVal nYear as Long) as Long

    Function = cmDaysFromGregorian(cCalendarClass.DECEMBER, _
                                   31, _
                                   nYear)

End Function
' ========================================================================================
' Return the Gregorian month,day, and year from a days date
' ========================================================================================
Private Sub cCalendar.cmGregorianFromDays (ByVal nDays as Long, _
                                           ByRef nMonth as Short, _
                                           ByRef nDay as Short, _
                                           ByRef nYear as Long)

Dim nPriorDays     as Long

' Calculate Year

    nYear = cmGregorianYearFromDays(nDays)

' Calculate Prior Days

    nPriorDays = nDays _
               - cmDaysFromGregorian(cCalendarClass.JANUARY, _
                                     1, _
                                     nYear)

' Adjust for assumption in above calculation
' that Feb always has 30 days

    Select Case nDays

        Case Is < cmDaysFromGregorian(cCalendarClass.MARCH, _
                                   1, _
                                   nYear)
    Case Else

        nPriorDays = nPriorDays + 1

        nPriorDays = nPriorDays _
                   + IIf(cmGregorianLeapYear(nYear) = True,0,1)

    End Select
'
' Calculate Month
'
    nMonth = cmFloor((12 * nPriorDays + 373) / 367)
'
' Calculate Day
'
    nDay = nDays _
         - cmDaysFromGregorian(nMonth,1,nYear) _
         + 1

End Sub
' ========================================================================================
' Given a Days date, return the gregorian year
' ========================================================================================
Private Function cCalendar.cmGregorianYearFromDays (ByVal nDays as Long) as Long

' 146097 represents the last day of a leap year of a 400 year cycle
' 1461 represents the last day of a 4 year cycle
' 36524 is the average length in days of one gregorian century

Dim nCenturies400     as Long
Dim nCenturies100     as Long
Dim nFourYearCycles   as Long
Dim nYears            as Long
Dim nD1               as Long
Dim nD2               as Long
Dim nD3               as Long
Dim nYear             as Long

    nDays = nDays - 1

' Number of Leap Year Centuries

    nCenturies400 = cmFloor(nDays / 146097)

' Number of Centuries

    nD1 = cmMod(nDays,146097)
    nCenturies100 = cmFloor(nD1 / 36524)
    nD2 = cmMod(nD1,36524)

' Number of 4 year cycles

    nFourYearCycles = cmFloor(nD2 / 1461)

' Number of Years

    nD3 = cmMod(nD2,1461)
    nYears = cmFloor(nD3 / 365)
    nYear = (400 * nCenturies400) _
          + (100 * nCenturies100) _
          + (4 * nFourYearCycles) _
          + nYears

' Adjustment for leap years past

' If nCenturies = 4 or nYears = 4 then we need to increment the year returned

    If nCenturies100 = 4 Then

    Else

        If nYears = 4 Then

        Else

            nYear = nYear + 1

        End If

    End If

    Function = nYear

End Function
' ========================================================================================
' Given a Gregorian month, day, and year, return the days date version.
' The days date represents the number of days since Jan 1, 1.
' ========================================================================================
Private Function cCalendar.cmDaysFromGregorian (ByVal nMonth as Long, _
                                                ByVal nDay as Long, _
                                                ByVal nYear as Long) as Long
                                        
Dim nDaysDate         as  Long
Dim nPriorYear        as  Long

    nMonth = Abs(nMonth)
    nDay = Abs(nDay)
    nPriorYear = nYear - 1

' Add:
'
' Number of days in prior years
' Number of days in prior months of current year
' Number of days in current month
'
' Assumption at this point is that Feb has 30 days

    nDaysDate = cCalendarClass.GREGORIAN_EPOCH - 1 _
              + (365 * nPriorYear) _                                 ' Days in prior years
              + cmFloor(nPriorYear / 4) _                            ' Leap Year days
              - cmFloor(nPriorYear / 100) _                          ' Century Years Adjust
              + cmFloor(nPriorYear / 400) _                          ' Century Years Adjust
              + cmFloor(((367 * nMonth) - 362) / 12)                 ' Days in prior months this year

' Adjust for assumption that Feb has 30 days

    Select Case nMonth

        Case Is < 3

        Case Else

            nDaysDate = nDaysDate - IIf(cmGregorianLeapYear(nYear) = True,1,2)

    End Select

    Function = nDaysDate + nDay

End Function
' ========================================================================================
' Determine if a Gregorian year is a leap year
' ========================================================================================
Private Function cCalendar.cmGregorianLeapYear (ByVal nYear as Long) as BOOLEAN

' Year is a leap year if evenly divisible by 4
' and if Year is a century year (ends with 00)
' it is also evenly divisible by 400

Dim nLeapYear                as BOOLEAN

    nLeapYear = False

    Select Case cmMod(nYear,4)

        Case Is <> 0

        Case Else

            Select Case cmMod(nYear,400)

                Case 100

                Case 200

                Case 300

                Case Else

                    nLeapYear = True

            End Select

    End Select

    Function = nLeapYear

End Function
' ========================================================================================
' Calculate the Gregorian day of the week
' ========================================================================================
Private Function cCalendar.cmGregorianWeekDay (ByVal nDays as Long) as Short

' Sun = 0
' Mon = 1
' Tue = 2
' Wed = 3
' Thu = 4
' Fri = 5
' Sat = 6

    Function = Abs(cmFloor(cmMod(nDays,7)))
    
End Function
' ========================================================================================
' Find the date of the first weekday on or after nDays
' ========================================================================================
Private Function cCalendar.cmFirstWeekDay (ByVal nWeekDay as Short, _
                                           ByVal nDays as Long) as Long

    Function = cmNthWeekDay(1, _
                            nWeekDay, _
                            nDays)

End Function
' ========================================================================================
' Find the date of the second weekday on or after nDays
' ========================================================================================
Private Function cCalendar.cmSecondWeekDay (ByVal nWeekDay as Short, _
                                            ByVal nDays as Long) as Long

    Function = cmNthWeekDay(2, _
                            nWeekDay, _
                            nDays)

End Function
' ========================================================================================
' Find the date of the third weekday on or after nDays
' ========================================================================================
Private Function cCalendar.cmThirdWeekDay (ByVal nWeekDay as Short, _
                                           ByVal nDays as Long) as Long

    Function = cmNthWeekDay(3, _
                            nWeekDay, _
                            nDays)

End Function
' ========================================================================================
' Find the date of the fourth weekday on or after nDays
' ========================================================================================
Private Function cCalendar.cmFourthWeekDay (ByVal nWeekDay as Short, _
                                            ByVal nDays as Long) as Long

    Function = cmNthWeekDay(4, _
                            nWeekDay, _
                            nDays)

End Function
' ========================================================================================
' Find the date of the first weekday on or before nDays
' ========================================================================================
Private Function cCalendar.cmLastWeekDay (ByVal nWeekDay as Short, _
                                          ByVal nDays as Long) as Long

    Function = cmNthWeekDay(-1, _
                            nWeekDay, _
                            nDays)

End Function
' ========================================================================================
' Find the nth occurrence of a weekday based on nDays
' ========================================================================================
Private Function cCalendar.cmNthWeekDay (ByVal nNthDay as Short, _
                                         ByVal nWeekDay as Short, _
                                         ByVal nDays as Long) as Long

    Select Case nNthDay

    Case Is > 0

        Function =  7 * nNthDay + cmWeekDayBefore(nWeekDay,nDays)

    Case Else

       Function =  7 * nNthDay + cmWeekDayAfter(nWeekDay,nDays)

    End Select

End Function
' ========================================================================================
' Calculate Weekday before nDays
' ========================================================================================
Private Function cCalendar.cmWeekDayBefore (ByVal nWeekDay as Short, _
                                            ByVal nDays as Long) as Long

    Function = cmWeekDayOnOrBefore(nWeekDay, _
                                   nDays - 1)

End Function
' ========================================================================================
' Calculate Weekday after nDays
' ========================================================================================
Private Function cCalendar.cmWeekDayAfter (ByVal nWeekDay as Short, _
                                           ByVal nDays as Long) as Long

    Function = cmWeekDayOnOrBefore(nWeekDay, _
                                   nDays + 7)

End Function
' ========================================================================================
' Calculate Weekday nearest to nDays
' ========================================================================================
Private Function cCalendar.cmWeekDayNearest (ByVal nWeekday as Short, _
                                             ByVal nDays as Long) as Long

    Function = cmWeekDayOnOrBefore(nWeekDay, _
                                   nDays + 3)

End Function
' ========================================================================================
' Calculate Weekday on or after nDays
' ========================================================================================
Private Function cCalendar.cmWeekDayOnOrAfter (ByVal nWeekDay as Short, _
                                               ByVal nDays as Long) as Long

    Function = cmWeekDayOnOrBefore(nWeekDay, _
                                   nDays + 6)
End Function
' ========================================================================================
' Calculate Weekday on or before nDays
' ========================================================================================
Private Function cCalendar.cmWeekDayOnOrBefore (ByVal nWeekDay as Short, _
                                                ByVal nDays as Long) as Long

    Function = nDays - _
               cmGregorianWeekDay(nDays - nWeekDay)
End Function

' ########################################################################################
' Chinese Support
' ########################################################################################

' ========================================================================================
' Chinese Marriage Auguries
' ========================================================================================
Private Function cCalendar.cmChineseYearMarriageAuguries (ByVal nCycle as Short, _
                                                          ByVal nYear as Long, _
                                                          ByVal nCountry as Short) as Short

' Chinese years that do not contain the minor term Beginning of Spring (Lichun) are widow years
' Chinese years that contain Beginning of Spring at both the beginning and the end
'    of the year are double bright years
' Chinese years missing the first Beginning of Spring but contain it at the end
'    of the year are blind years
' Chinese years that contain the first Beginning of Spring but do not have it at the end
'    of the year are bright years

' Chinese tradition deems it unlucky to be married in a widow year

Dim nNewYear        as Long
Dim nNextCycle      as Short
Dim nNextYear       as Long
Dim nNextNewYear    as Long
Dim nFirstTerm      as Long
Dim nNextTerm       as Long
Dim nAugury         as Short

    nNewYear = cmDaysFromChinese(nCycle,nYear,1,False,1,nCountry)
    nNextCycle = nCycle + IIf(nYear = 60,1,0)
    nNextYear = 1 + IIf(nYear <> 60,nYear,0)
    nNextNewYear = cmDaysFromChinese(nNextCycle,nNextYear,1,False,1,nCountry)
    nFirstTerm = cmCurrentMinorSolarTerm(nNewYear,nCountry)
    nNextTerm = cmCurrentMinorSolarTerm(nNextNewYear,nCountry)
    
    Select Case nFirstTerm
    
       Case 1
       
          nAugury = IIf(nNextTerm = 12,cCalendarClass.CHINESE_WIDOW_YEAR,cCalendarClass.CHINESE_BLIND_YEAR)   
       
       Case Else

          nAugury = IIf(nNextTerm = 12,cCalendarClass.CHINESE_BRIGHT_YEAR,cCalendarClass.CHINESE_DOUBLE_BRIGHT_YEAR)       
       
    End Select
    
    Function = nAugury
    
End Function
' ========================================================================================
' Chinese Year Name
' ========================================================================================
Private Function cCalendar.cmChineseYearName (ByVal nYear as Long) as Short

    Function = cmChineseSexagesimalName(nYear)
    
End Function
' ========================================================================================
' Chinese Month Name
' ========================================================================================
Private Function cCalendar.cmChineseMonthName (ByVal nMonth as Short, _
                                               ByVal nYear as Long) as Short
                                      
Dim nElaspedMonths  as Long

    nElaspedMonths = 12 * (nYear - 1) + nMonth - 1

    Function = cmChineseSexagesimalName(nElaspedMonths - cCalendarClass.CHINESE_MONTH_NAME_EPOCH)
    
End Function
' ========================================================================================
' Find the nth name of the sexagenary cycle of year names
' ========================================================================================
Private Function cCalendar.cmChineseSexagesimalName(ByVal nYear as Long) as Short

    Function = cmAMod(cmAMod(nYear,12),10)

End Function
' ========================================================================================
' Given a Chinese date, return a days date
' ========================================================================================
Private Function cCalendar.cmDaysFromChinese (ByVal nCycle as Short, _
                                              ByVal nYear as Long, _
                                              ByVal nMonth as Short, _
                                              ByVal bLeapMonth as BOOLEAN, _
                                              ByVal nDay as Short, _
                                              ByVal nCountry as Short) as Long

Dim nNewYear               as Long
Dim nMidYear               as Long
Dim nNextNewMoon           as Long
Dim nWorkCycle             as Short
Dim nWorkYear              as Long
Dim nWorkMonth             as Short
Dim nWorkLeapMonth         as BOOLEAN
Dim nWorkDay               as Short
Dim nPriorNewMoon          as Long

    nMidYear = cmFloor(cCalendarClass.CHINESE_EPOCH + ((nCycle - 1) * 60 + nYear - 1 + .5) * 365.242189)
    nNewYear = cmChineseNewYearOnOrBefore(nMidYear,nCountry)
    nNextNewMoon = cmChineseNewMoonOnOrAfter(nNewYear + ((nMonth - 1) * 29),nCountry)
    cmChineseFromDays(nNextNewMoon,nWorkCycle,nWorkYear,nWorkMonth,nWorkLeapMonth,nWorkDay,nCountry)

    If nWorkMonth = nMonth AndAlso nWorkLeapMonth = bLeapMonth Then

        nPriorNewMoon = nNextNewMoon

    Else

        nPriorNewMoon = cmChineseNewMoonOnOrAfter(nNextNewMoon + 1,nCountry)

    End If

    Function = nPriorNewMoon + nDay - 1

End Function
' ========================================================================================
' Return Chinese dates occurring in a given Gregorian year
' Chinese holidays never occur in a leap month so we can default FALSE
' ========================================================================================
Private Function cCalendar.cmChineseDateCalculation (ByVal nRuleClass as Short, _
                                                     ByVal nMonth as Short, _
                                                     ByVal nDay as Short, _
                                                     ByVal nGregorianYear as Long, _
                                                     ByVal nRule as Short, _
                                                     ByRef nCalcDays as Long) as BOOLEAN

Dim nJan1             as Long
Dim nDec31            as Long
Dim nCycle            as Short
Dim nChineseYear      as Long
Dim nCountry          as Short

' Select country for calendar

    Select Case nRuleClass

    Case cCalendarClass.VIETNAMESE_RULES

        nCountry = cCalendarClass.VIETNAMESE

    Case cCalendarClass.KOREAN_RULES

        nCountry = cCalendarClass.KOREAN

    Case cCalendarClass.JAPANESE_RULES

        nCountry = cCalendarClass.KOREAN

    Case Else

        nCountry = cCalendarClass.CHINESE

    End Select

' If we have a special rule, we can guarantee finding it within a gregorian year

    Select Case nRule

    Case cCalendarClass.CHINESE_WINTERSOLSTICE

        nCalcDays = cmChineseWinterSolsticeOnOrBefore(cmDaysFromGregorian(cCalendarClass.DECEMBER,30,nGregorianYear),nCountry)
        
        Function = True

    Case cCalendarClass.CHINESE_QINGMING

        nCalcDays = cmFloor(cmMinorSolarTermOnOrAfter(cmDaysFromGregorian(cCalendarClass.MARCH,30,nGregorianYear),nCountry))
        
        Function = True

    Case Else

        cmChineseCycleAndYear(nGregorianYear,nCycle,nChineseYear)
        cmGregorianYearRange(nGregorianYear,nJan1,nDec31)

' Calculate the Chinese Cycle and Year

        nCalcDays = cmDaysFromChinese(nCycle,nChineseYear,nMonth,False,nDay,nCountry)

        Select Case nCalcDays

' Check if Chinese date occurs during nGregorianYear

        Case nJan1 To nDec31
        
             Function = True

        Case Else

            cmChineseCycleAndYear(nGregorianYear - 1,nCycle,nChineseYear)

            nCalcDays = cmDaysFromChinese(nCycle,nChineseYear,nMonth,False,nDay,cCalendarClass.CHINESE)

' Check if Chinese date occurs at all during nGregorianYear

            Select Case nCalcDays

            Case nJan1 To nDec31

                Function = True

            Case Else

                Function = False

            End Select

        End Select

    End Select

End Function
' ========================================================================================
' Given days date nDays, return the Chinese equivalent
' ========================================================================================
Private Sub cCalendar.cmChineseFromDays (ByVal nDays as Long, _
                                         ByRef nCycle as Short, _
                                         ByRef nYear as Long, _
                                         ByRef nMonth as Short, _
                                         ByRef bLeapMonth as BOOLEAN, _
                                         ByRef nDay as Short, _
                                         ByVal nCountry as Short)

Dim nS1                    as Long
Dim nS2                    as Long
Dim nNextM11               as Long
Dim nM                     as Long
Dim nM12                   as Long
Dim nLeapYear              as BOOLEAN
Dim nElaspedYears          as Long

    nS1 = cmChineseWinterSolsticeOnOrBefore(nDays,nCountry)
    nM12 = cmChineseNewMoonOnOrAfter(nS1 + 1,nCountry)
    nM = cmChineseNewMoonBefore(nDays + 1,nCountry)
    nS2 = cmChineseWinterSolsticeOnOrBefore(nS1 + 370,nCountry)
    nNextM11 = cmChineseNewMoonBefore(nS2 + 1,nCountry)
    nLeapYear = IIf(cmRound((nNextM11 - nM12) / cCalendarClass.MeanSynodicMonth) = 12,True,False)
    nMonth = cmRound((nM - nM12) / cCalendarClass.MeanSynodicMonth)

    If nLeapYear = True AndAlso cmChinesePriorLeapMonth(nM12,nM,nCountry) = True Then

        nMonth = nMonth - 1

    End If

    nMonth = cmAMod(nMonth,12)

    If nLeapYear = True AndAlso cmChineseNoMajorSolarTerm(nM,nCountry) = True AndAlso _
            cmChinesePriorLeapMonth(nM12,cmChineseNewMoonBefore(nM,nCountry),nCountry) = False Then

        bLeapMonth = True

    Else

        bLeapMonth = False

    End If

    nElaspedYears = cmFloor(1.5 - (nMonth / 12) + ((nDays - cCalendarClass.CHINESE_EPOCH) / 365.242189))

    nCycle = cmFloor((nElaspedYears - 1) / 60) + 1

    nYear = cmAMod(nElaspedYears,60)

    nDay = nDays - nM + 1

End Sub
' ========================================================================================
' Given a Gregorian year, return the Chinese cycle and year equivalent
' ========================================================================================
Private Sub cCalendar.cmChineseCycleAndYear (ByVal nGregorianYear as Long, _
                                             ByRef nCycle as Short, _
                                             ByRef nChineseYear as Long)

Dim nElaspedYears     as Long

    nElaspedYears = nGregorianYear + 2637

    nCycle = cmFloor((1 / 60) * (nElaspedYears - 1)) + 1

    nChineseYear = cmAMod(nElaspedYears,60)

End Sub
' ========================================================================================
' Find Chinese new year on or before nDays
' ========================================================================================
Private Function cCalendar.cmChineseNewYearOnOrBefore (ByVal nDays as Long, _
                                                       ByVal nCountry as Short) as Long

Dim nNewYear       as Long

    nNewYear = cmChineseNewYearInSui(nDays,nCountry)

    Select Case nDays

    Case Is >= nNewYear

    Case Else

        nNewYear = cmChineseNewYearInSui(nDays - 180,nCountry)

    End Select

    Function = nNewYear

End Function
' ========================================================================================
' Return first day of Chinese year for the sui containing nDays
' ========================================================================================
Private Function cCalendar.cmChineseNewYearInSui (ByVal nDays as Long, _
                                                  ByVal nCountry as Short) as Long

Dim nS1            as Long
Dim nS2            as Long
Dim nM12           as Long
Dim nM13           as Long
Dim nNextM11       as Long
Dim nNewYear       as Long

    nS1 = cmChineseWinterSolsticeOnOrBefore(nDays,nCountry)
    nS2 = cmChineseWinterSolsticeOnOrBefore(nS1 + 370,nCountry)
    nM12 = cmChineseNewMoonOnOrAfter(nS1 + 1,nCountry)
    nM13 = cmChineseNewMoonOnOrAfter(nM12 + 1,nCountry)
    nNextM11 = cmChineseNewMoonBefore(nS2 + 1,nCountry)
    nNewYear = cmChineseNewMoonOnOrAfter(nM13 + 1,nCountry)

    If cmRound((nNextM11 - nM12) / cCalendarClass.MeanSynodicMonth) = 12 Then

        If cmChineseNoMajorSolarTerm(nM12,nCountry) = True OrElse cmChineseNoMajorSolarTerm(nM13,nCountry) = True Then

        Else

            nNewYear = nM13

        End If

    Else

        nNewYear = nM13

    End If

    Function =  nNewYear

End Function
' ========================================================================================
' Returns True if there is a Chinese leap month at or after lunar
' month nMPrime and at or before lunar month nM
' ========================================================================================
Private Function cCalendar.cmChinesePriorLeapMonth (ByVal nMPrime as Long, _
                                                    ByVal nM as Long, _
                                                    ByVal nCountry as Short) as BOOLEAN

Dim nReturn            as BOOLEAN
Dim nLoop              as BOOLEAN

    nReturn = False
    nLoop = True

    Select Case nM

    Case Is >= nMPrime

        While nLoop = True

            nReturn = cmChineseNoMajorSolarTerm(nM,nCountry)

            Select Case nReturn

            Case True

                nLoop = False

            Case Else

                nM = cmChineseNewMoonBefore(nM,nCountry)
                nLoop = IIf(nM >= nMPrime,True,False)

            End Select

        Wend

    End Select

    Function = nReturn

End Function
' ========================================================================================
' Returns TRUE if Chinese lunar month starting on nDays
' has no major solar term FALSE otherwise
' ========================================================================================
Private Function cCalendar.cmChineseNoMajorSolarTerm (ByVal nDays as Long, _
                                                      ByVal nCountry as Short) as BOOLEAN

    Function = IIf(cmCurrentMajorSolarTerm(nDays,nCountry) = cmCurrentMajorSolarTerm(cmChineseNewMoonOnOrAfter(nDays + 1,nCountry),nCountry),True,False)

End Function
' ========================================================================================
' Date of the first new moon before nMoment
' ========================================================================================
Private Function cCalendar.cmChineseNewMoonBefore (ByVal nMoment as Double, _
                                                   ByVal nCountry as Short) as Long

Dim nNewMoon   as Double

    nNewMoon = cmNewMoonBefore(cmMidnightInChina(nMoment,nCountry))

    Function = cmFloor(cmStandardFromUniversal(nNewMoon,cmChineseLocation(nNewMoon,nCountry)))

End Function
' ========================================================================================
' Date of the first new moon on or after nMoment
' ========================================================================================
Private Function cCalendar.cmChineseNewMoonOnOrAfter (ByVal nMoment as Double, _
                                                      ByVal nCountry as Short) as Long

Dim nNewMoon   as Double

    nNewMoon = cmNewMoonAfter(cmMidnightInChina(nMoment,nCountry))

    Function = cmFloor(cmStandardFromUniversal(nNewMoon,cmChineseLocation(nNewMoon,nCountry)))

End Function
' ========================================================================================
' Date of Winter Solstice on or before nMoment
' ========================================================================================
Private Function cCalendar.cmChineseWinterSolsticeOnOrBefore (ByVal nDays as Long, _
                                                              ByVal nCountry as Short) as Long

Dim nSolstice      as Double
Dim nLoop          as BOOLEAN

    nSolstice = cmEstimatePriorSolarLongitude(cmMidnightInChina(nDays + 1,nCountry),cCalendarClass.WINTER)
    nLoop = True

    While nLoop = True

        If cCalendarClass.WINTER > cmSolarLongitude(cmMidnightInChina(cmFloor(nSolstice) + 1,nCountry)) Then _

            nSolstice = nSolstice + 1

        Else

            nLoop = False

        End If

    Wend

    Function = cmFloor(nSolstice)

End Function
' ========================================================================================
' Date Chinese minor solar term (jieqi) on or after nMoment
' ========================================================================================
Private Function cCalendar.cmMinorSolarTermOnOrAfter (ByVal nMoment as Double, _
                                                      ByVal nCountry as Short) as Double

Dim nSolarTerm     as Long

    nSolarTerm = cmCalcDegrees(30 * cmCeiling((cmSolarLongitude(cmMidnightInChina(nMoment,nCountry)) - 15) / 30) + 15)

    Function = cmChineseSolarLongitudeOnOrAfter(nMoment,nSolarTerm,nCountry)

End Function
' ========================================================================================
' Last Chinese minor solar term (jieqi) index before nMoment
' ========================================================================================
Private Function cCalendar.cmCurrentMinorSolarTerm (ByVal nMoment as Double, _
                                                    ByVal nCountry as Short) as Long

Dim nLongitude     as Double

    nLongitude = cmSolarLongitude(cmUniversalFromStandard(nMoment,cmChineseLocation(nMoment,nCountry)))

    Function = cmAMod(3 + cmFloor((nLongitude - 15) / 30),12)

End Function
' ========================================================================================
' Date of the first Chinese major solar term (zhongqi) on or after nMoment. Major solar
' terms begin when the sun's longitude is a multiple of 30 degrees
' ========================================================================================
Private Function cCalendar.cmMajorSolarTermOnOrAfter (ByVal nMoment as Double, _
                                                      ByVal nCountry as Short) as Double

Dim nSolarTerm     as Long

    nSolarTerm = cmCalcDegrees(30 * cmCeiling(cmSolarLongitude(cmMidnightInChina(nMoment,nCountry)) / 30))

    Function = cmChineseSolarLongitudeOnOrAfter(nMoment,nSolarTerm,nCountry)

End Function
' ========================================================================================
' Chinese Midnight in Universal Time
' ========================================================================================
Private Function cCalendar.cmMidnightInChina (ByVal nMoment as Double, _
                                              ByVal nCountry as Short) as Double

    Function = cmUniversalFromStandard(nMoment,cmChineseLocation(nMoment,nCountry))

End Function
' ========================================================================================
' Moment of the first date on or after nMoment when the solar longitude is a multiple of
' nSolarTerm degrees
' ========================================================================================
Private Function cCalendar.cmChineseSolarLongitudeOnOrAfter (ByVal nMoment as Double, _
                                                             ByVal nSolarTerm as Long, _
                                                             ByVal nCountry as Short) as Double

Dim nLongitude     as Double
Dim nZone          as Double

    nZone = cmChineseLocation(nMoment,nCountry)

    nLongitude = cmSolarLongitudeAfter(cmUniversalFromStandard(nMoment,nZone),nSolarTerm)

    Function = cmStandardFromUniversal(nLongitude,nZone)

End Function
' ========================================================================================
' Last Chinese major solar term (zhongqi) index before nMoment
' ========================================================================================
Private Function cCalendar.cmCurrentMajorSolarTerm (ByVal nMoment as Double, _
                                                    ByVal nCountry as Short) as Long

Dim nLongitude         as Double

    nLongitude = cmSolarLongitude(cmUniversalFromStandard(nMoment,cmChineseLocation(nMoment,nCountry)))

    Function = cmAMod(2 + cmFloor(nLongitude / 30),12)

End Function
' ========================================================================================
' Determine zone hours based on country option
' ========================================================================================
Private Function cCalendar.cmChineseLocation (ByVal nMoment as Double, _
                                              ByVal nCountry as Short) as Double

Dim nYear      as Long

    nYear = cmGregorianYearFromDays(cmFloor(nMoment))

    Select Case nCountry

    Case cCalendarClass.VIETNAMESE

        Select Case cmFloor(nMoment)

            Case Is < cmDaysFromGregorian(cCalendarClass.JANUARY,1,1968)

                Function = 8

            Case Else

                Function = 7

        End Select

    Case cCalendarClass.KOREAN

        Select Case cmFloor(nMoment)

            Case Is < cmDaysFromGregorian(cCalendarClass.APRIL,1,1908)

                Function = 8.4644444444

            Case Is < cmDaysFromGregorian(cCalendarClass.JANUARY,1,1912)

                Function = 8.5

            Case Is < cmDaysFromGregorian(cCalendarClass.MARCH,21,1954)

                Function = 9

            Case Is < cmDaysFromGregorian(cCalendarClass.AUGUST,10,1961)

                Function = 8.5

            Case Else

                Function = 9

        End Select

    Case cCalendarClass.JAPANESE

        Function = IIf(nYear < 1888,9.3177777778,9)

    Case Else

        Function = IIf(nYear < 1929,1397 / 180,8)

    End Select

End Function

' ########################################################################################
' 13 Month Rolling History Support
' ########################################################################################

' ========================================================================================
' Add one month to running summary total
' ========================================================================================
Private Sub cCalendar.cmSumOneMonth (ByVal nMonth as Long, _
                                     uHistory as HISTORY_MONTHS, _
                                     ByRef nSummary as Double)

    nSummary = nSummary + uHistory.Month(nMonth)

End Sub
' ========================================================================================
' Clear Summary totals
' ========================================================================================
Private Sub cCalendar.cmClearSummary (arSummary() as Double)

Dim nIndex        as Long

    For nIndex = 0 To UBound(arSummary)

        arSummary(nIndex) = 0

    Next

End Sub
' ========================================================================================
' Save and Clear Current History
' ========================================================================================
Private Sub cCalendar.cmShiftHistory (ByVal nMonth as Short, _
                                      uHistory as HISTORY_MONTHS)

    uHistory.Month(nMonth) = uHistory.Month(0)
    uHistory.Month(0) = 0

End Sub

' ########################################################################################
' Julian Support
' ########################################################################################

' ========================================================================================
' Return the Julian month,day, and year from a days date
' ========================================================================================
Private Sub cCalendar.cmJulianFromDays(ByVal nDays as Long, _
                                       ByRef nMonth as Short, _
                                       ByRef nDay as Short, _
                                       ByRef nYear as Long)
                     
' 1461 represents the last day of a 4 year cycle

Dim nApprox     as Long
Dim nPriorDays  as Long
Dim nCorrection as Long

    nApprox = cmFloor((1 / 1461) * (4 * (nDays - cCalendarClass.JULIAN_EPOCH) + 1464))
    
    Select Case nApprox
    
        Case Is <= 0
        
            nYear = nApprox - 1
            
        Case Else
        
            nYear = nApprox
            
    End Select
    
    nPriorDays = nDays - cmDaysFromJulian(cCalendarClass.JANUARY,1,nYear)
    
    If nDays < cmDaysFromJulian(cCalendarClass.MARCH,1,nYear) Then
    
        nCorrection = 0
        
    Else
    
        If cmJulianLeapYear(nYear) = True Then
        
            nCorrection = 1
            
        Else
        
            nCorrection = 2
            
        End If
        
    End If 
       
    nMonth = cmFloor((1 / 367) * (12 * (nPriorDays + nCorrection) + 373))
    
    nDay = nDays - cmDaysFromJulian(nMonth,1,nYear) + 1

End Sub
' ========================================================================================
' Calculate days date from Julian date
' ========================================================================================
Private Function cCalendar.cmDaysFromJulian (ByVal nMonth as Short, _
                                             ByVal nDay as Short, _
                                             ByVal nYear as Long) as Long

Dim nMonthAdjust     as Long
Dim nDays            as Long

    Select Case nMonth

    Case Is <= 2

        nMonthAdjust = 0

    Case Else

        nMonthAdjust = IIf(cmJulianLeapYear(nYear) = True,-1,-2)

    End Select

' Year has to be adjusted since year zero is not valid

    nYear = nYear + IIf(nYear < 1,1,0)

    Function = (cCalendarClass.JULIAN_EPOCH - 1 + 365 * (nYear - 1)) _
             + cmFloor((nYear - 1) / 4) _
             + cmFloor((nMonth * 367 - 362) / 12) _
             + nMonthAdjust _
             + nDay

End Function
' ========================================================================================
' Determine if nYear is a Julian leap year
' ========================================================================================
Private Function cCalendar.cmJulianLeapYear (ByVal nYear as Long) as BOOLEAN

Dim nReturn            as BOOLEAN

    nReturn = False

    Select Case nYear

    Case Is > 0

        Select Case cmFloor(cmMod(nYear,4))

        Case 0

            nReturn = True

        End Select

    Case Else

        Select Case cmFloor(cmMod(nYear,4))

        Case 3

            nReturn = True

        End Select

    End Select

   Function = nReturn

End Function

' ########################################################################################
' Astronomy support
' ########################################################################################

' ========================================================================================
' Lunar Illumination
' ========================================================================================
Private Function cCalendar.cmLunarIllumination(byVal nMoment as Double) as Double

' Lunar Illumination 0.00 - 1.00 at universal moment

Dim nSolarDistance  as Double
Dim nLunarDistance  as Double
Dim nLunarLatitude  as Double
Dim nLunarLongitude as Double
Dim nSolarLongitude as Double
Dim nLunarPhase     as Double

    nSolarDistance = cmSolarDistance(nMoment) * 149597870.691
    nLunarDistance = cmLunarDistance(nMoment) / 1000
    nLunarLatitude = cmLunarLatitude(nMoment)
    nLunarLongitude = cmLunarLongitude(nMoment)
    nSolarLongitude = cmSolarLongitude(nMoment)
    nLunarPhase = cmCoSineDegrees(nLunarLatitude) _
                * cmCoSineDegrees(nLunarLongitude - nSolarLongitude)
    nLunarPhase = cmArcCoSineDegrees(nLunarPhase)

    nLunarPhase = (nSolarDistance * cmSinDegrees(nLunarPhase)) _
                / (nLunarDistance - (nSolarDistance * cmCoSineDegrees(nLunarPhase)))

    nLunarPhase = Atn(nLunarPhase)
    nLunarPhase = cmRadiansToDegrees(nLunarPhase)

' Normalize to maximum lunar visible half circle

    nLunarPhase = cmMod(nLunarPhase,180)

    Function = (1 + cmCosineDegrees(nLunarPhase)) / 2

End Function
' ========================================================================================
' Moonrise and Moonset Local Times for one day
' ========================================================================================
Private Sub cCalendar.cmMoonRiseAndSet (ByVal nSerial as LongInt, _
                                        ByVal bType as BOOLEAN, _
                                        ByRef uLocale as LOCATION_LOCALE, _
                                        arLunarTimes() as LUNAR_RISE_AND_SET)

' arLunarTimes contains array entries per event

' Geocentric times are calculated from the center of the earth not corrected for
' parallax and refraction. Many published times are geocentric calculations,
' including the USNO tables. Topocentric times are calculated from the earth
' surface and are corrected for parallax and refraction.

' The times are found by looking for hours when the altitude of the moon
' crosses the horizon (the sign changes) and then a bisection search is found
' within the hour to find the time to the nearest second before the found event
' time is rounded to the nearest minute.

Dim nLastAltitude         as Short
Dim nCurrentAltitude      as Short
Dim bDaylightSavings      as BOOLEAN
Dim nSearchHour           as LongInt
Dim nSearchHourFrom       as LongInt
Dim nSearchHourTo         as LongInt
Dim nSearchMinute         as LongInt
Dim bEventType            as BOOLEAN    ' False=Moonset, True=Moonrise
Dim nEventTime            as LongInt
Dim bSearch               as BOOLEAN 
Dim nSearchAltitude       as Double                  

' Save the current locale DST switch to be restored later

    bDaylightSavings = uLocale.bDaylightLightSavingsActive

' Since nLunarDay will be assumed to be a UTC date, well search the prior and next
' days to be sure to find events on nSerial in local time.

    If bType = cCalendarClass.GEOCENTRIC Then
       nLastAltitude = cmSignum(cmGeocentricLunarAltitude(cmSerialToMoment(nSerial - cCalendarClass.ONE_DAY),uLocale))
    Else
       nLastAltitude = cmSignum(cmTopocentricLunarAltitude(cmSerialToMoment(nSerial - cCalendarClass.ONE_DAY),uLocale))
    End If 
    
    For nSearchHour = nSerial - cCalendarClass.ONE_DAY + cCalendarClass.ONE_HOUR To nSerial + cCalendarClass.ONE_DAY * 2 Step cCalendarClass.ONE_HOUR

' Calculate current hour lunar altitude

        If bType = cCalendarClass.GEOCENTRIC Then
           nCurrentAltitude = cmSignum(cmGeocentricLunarAltitude(cmSerialToMoment(nSearchHour),uLocale))
        Else
           nCurrentAltitude = cmSignum(cmTopocentricLunarAltitude(cmSerialToMoment(nSearchHour),uLocale))
        End If 

' Check for rise or set event

        If nCurrentAltitude <> nLastAltitude Then
        
           If nCurrentAltitude < 0 Then

' Moonset    
              bEventType = cCalendarClass.MOONSET
              nLastAltitude = nCurrentAltitude

           Else

'Moonrise
              bEventType = cCalendarClass.MOONRISE
              nLastAltitude = nCurrentAltitude

           End If

' Search with the hour to find the event

           nSearchHourFrom = nSearchHour - cCalendarClass.ONE_HOUR
           nSearchHourTo = nSearchHour
           nSearchMinute = nSearchHourFrom + ((nSearchHourTo - nSearchHourFrom) / 2)
           bSearch = True
           
           While bSearch = True

              If bType = cCalendarClass.GEOCENTRIC Then
  
                nSearchAltitude = cmGeocentricLunarAltitude(cmSerialToMoment(nSearchMinute),uLocale)

              Else

                nSearchAltitude = cmTopocentricLunarAltitude(cmSerialToMoment(nSearchMinute),uLocale)

             End If

             Select Case nSearchAltitude
    
                Case Is < 0

                   If bEventType = cCalendarClass.MOONRISE Then

' Lower range for moonrise

                      nSearchHourFrom = nSearchMinute

                   Else

' Upper range for moonset

                      nSearchHourTo = nSearchMinute

                   End If
       
                Case Is > 0

                   If bEventType = cCalendarClass.MOONRISE Then

' Upper range for moonrise

                      nSearchHourTo = nSearchMinute

                   Else

' Lower range for moonset

                      nSearchHourFrom = nSearchMinute

                   End If

                Case Else

' Exact 0 altitude
       
                   bSearch = False
       
             End Select
     
             nSearchMinute = nSearchHourFrom + ((nSearchHourTo - nSearchHourFrom) / 2) 
     
             If nSearchHourTo - nSearchHourFrom <= cCalendarClass.ONE_SECOND Then
     
                bSearch = False
        
             End If

           Wend
           
Dim uGreg as GREGORIAN_DATE

' Round to nearest minute

           nSearchMinute = cmFloor((nSearchMinute + (30 * cCalendarClass.ONE_SECOND)) _
                         / cCalendarClass.ONE_MINUTE) * cCalendarClass.ONE_MINUTE

' Convert to local time with daylight savings applied, if active

           nSearchMinute = cmDaylightSavings(cmMomentToSerial(cmStandardFromUniversal(cmSerialToMoment(nSearchMinute),uLocale.Zone)),uLocale)
 
' Check if it for the day requested

           Select Case cmFloor(cmSerialToMoment(nSearchMinute))

              Case cmFloor(cmSerialToMoment(nSerial))

                 ReDim Preserve arLunarTimes(UBound(arLunarTimes) + 1)
                 arLunarTimes(UBound(arLunarTimes)).LunarSerialTime = nSearchMinute
                 arLunarTimes(UBound(arLunarTimes)).DaylightSavings = uLocale.bDaylightLightSavingsActive
                 arLunarTimes(UBound(arLunarTimes)).RiseOrSet = bEventType
                    
              Case Is > cmFloor(cmSerialToMoment(nSerial))

                 Exit For
                    
              End Select
           
      End If
           
    Next

    uLocale.bDaylightLightSavingsActive = bDaylightSavings

End Sub
' ========================================================================================
' Topocentric Lunar Altitude
' ========================================================================================
Private Function cCalendar.cmTopocentricLunarAltitude (ByVal nMoment as Double, _
                                                       ByRef uLocale as LOCATION_LOCALE) as Double

' Correct geocentric altitude from earth center to surface
' and adjust for parallax and refraction

Dim nLunarAltitude    as Double

    nLunarAltitude = cmGeocentricLunarAltitude(nMoment,uLocale)
    Function = nLunarAltitude _
             - cmLunarParallax(nMoment,nLunarAltitude,uLocale.Latitude,uLocale.Longitude) _
             + cmSolarRefraction(uLocale.Elevation,uLocale.Latitude)

End Function
' ========================================================================================
' Lunar Parallax
' ========================================================================================
Private Function cCalendar.cmLunarParallax (ByVal nMoment as Double, _
                                            ByVal nLunarAltitude as Double, _
                                            ByVal nLatitude as Double, _
                                            ByVal nLongitude as Double) as Double

    Function = cmArcSinDegrees((cmEarthRadius(nLatitude) / cmLunarDistance(nMoment)) _
             * cmCoSineDegrees(nLunarAltitude))

End Function
' ========================================================================================
' Geocentric Altitude of the Moon above the horizon at UTC nMoment
' ========================================================================================
Private Function cCalendar.cmGeocentricLunarAltitude (ByVal nMoment as Double, _
                                                      ByRef uLocale as LOCATION_LOCALE) as Double


' Not corrected for parallax or refraction

Dim nObliquity                 as Double
Dim nLunarLongitude            as Double
Dim nLunarLatitude             as Double
Dim nLunarRightAscension       as Double
Dim nLunarDeclination          as Double
Dim nLocalSiderealHourAngle    as Double
Dim nAltitude                  as Double

    nObliquity = cmObliquity(cmJulianCenturies(nMoment))
    nLunarLongitude = cmLunarLongitude(nMoment)
    nLunarLatitude = cmLunarLatitude(nMoment)
    nLunarRightAscension = cmRightAscension(nObliquity,nLunarLatitude,nLunarLongitude)
    nLunarDeclination = cmDeclination(nObliquity,nLunarLatitude,nLunarLongitude)
    nLocalSiderealHourAngle = cmCalcDegrees(cmSiderealFromMoment(nMoment) + _
                                            uLocale.Longitude - nLunarRightAscension)
    nAltitude = cmArcSinDegrees((cmSinDegrees(uLocale.Latitude) * _
                      cmSinDegrees(nLunarDeclination)) + _
                      (cmCoSineDegrees(uLocale.Latitude) * _
                      cmCoSineDegrees(nLunarDeclination) * _
                      cmCoSineDegrees(nLocalSiderealHourAngle)))

    Function = cmCalcDegrees(nAltitude + 180) - 180

End Function
' ========================================================================================
' Distance between the centers of the Earth and Sun in Astronomical Units
' 1 AU = 149,597,870.691 kilometers or 92,955,807.267433 miles
' ========================================================================================
Private Function cCalendar.cmSolarDistance (ByVal nMoment as Double) as Double

Dim nEccentricity          as Double
Dim nSolarAnomaly          as Double
Dim nC                     as Double

    nC = cmJulianCenturies(nMoment)

' Eccentricity of Earth's orbit

    nEccentricity = cmEccentricityEarthOrbit(nC)
    nSolarAnomaly = cmCalcDegrees(357.5291092 + 35999.0502909 * nC - .0001537 * nC^2) _
                  + cmSolarEquationOfCenter(nC)

    Function = (1.000001018 * (1 - nEccentricity^2)) / (1 + nEccentricity * cmCoSineDegrees(nSolarAnomaly))

End Function
' ========================================================================================
' Equation of Center for the Sun
' ========================================================================================
Private Function cCalendar.cmSolarEquationOfCenter (ByVal nC as Double) as Double

Dim nMeanAnomaly   as Double
Dim nSinM          as Double
Dim nSin2M         as Double
Dim nSin3M         as Double

    nMeanAnomaly = (cCalendarClass.PI * cmSolarMeanAnomaly(nC)) / 180
    nSinM = Sin(nMeanAnomaly)
    nSin2M = Sin(nMeanAnomaly + nMeanAnomaly)
    nSin3M = Sin(nMeanAnomaly + nMeanAnomaly + nMeanAnomaly)

    Function = nSinM * (1.914602 - nC * (.004817 + .000014 * nC)) _
             + nSin2M * (.019993 - .000101 * nC) _
             + nSin3M * .000289

End Function
' ========================================================================================
' Calculate Sunrise in nZone time
' ========================================================================================
Private Function cCalendar.cmSunRise (ByVal nDays as Long, _
                                      ByVal nZone as Double, _
                                      ByVal nLatitude as Double, _
                                      ByVal nLongitude as Double, _
                                      ByVal nElevation as Double, _            ' In meters
                                      ByVal nDepression as Double, _          
                                      ByRef bBogus as BOOLEAN) as Double

    Function = cmDawn(nDays,nZone,nLatitude,nLongitude,nDepression + cmSolarRefraction(nElevation,nLatitude),bBogus)

End Function
' ========================================================================================
' Calculate Sunset in nZone time
' ========================================================================================
Private Function cCalendar.cmSunSet (ByVal nDays as Long, _
                                     ByVal nZone as Double, _
                                     ByVal nLatitude as Double, _
                                     ByVal nLongitude as Double, _
                                     ByVal nElevation as Double, _             ' In meters
                                     ByVal nDepression as Double, _
                                     ByRef bBogus as BOOLEAN) as Double

    Function = cmDusk(nDays,nZone,nLatitude,nLongitude,nDepression + cmSolarRefraction(nElevation,nLatitude),bBogus)

End Function
' ========================================================================================
' Calculate Dawn
' ========================================================================================
Private Function cCalendar.cmDawn (ByVal nDays as Long, _
                                   ByVal nZone as Double, _
                                   ByVal nLatitude as Double, _
                                   ByVal nLongitude as Double, _
                                   ByVal nDepression as Double, _
                                   ByRef bBogus as BOOLEAN) as Double

Dim nEvent     as Double

    nEvent = cmMomentOfDepression(nDays + .25, _
                                  nLatitude, _
                                  nLongitude, _
                                  nDepression, _
                                  cCalendarClass.MORNING, _
                                  bBogus)

    Function = cmStandardFromLocal(nEvent,nZone,nLongitude)

End Function
' ========================================================================================
' Calculate Dusk
' ========================================================================================
Private Function cCalendar.cmDusk (ByVal nDays as Long, _
                                   ByVal nZone as Double, _
                                   ByVal nLatitude as Double, _
                                   ByVal nLongitude as Double, _
                                   ByVal nDepression as Double, _
                                   ByRef bBogus as BOOLEAN) as Double

Dim nEvent     as Double

    nEvent = cmMomentOfDepression(nDays + .75, _
                                  nLatitude, _
                                  nLongitude, _
                                  nDepression, _
                                  cCalendarClass.EVENING, _
                                  bBogus)

    Function = cmStandardFromLocal(nEvent,nZone,nLongitude)

End Function
' ========================================================================================
' Find Moment when the Sun is at nDepression angle
' ========================================================================================
Private Function cCalendar.cmMomentOfDepression (ByVal nApprox as Double, _
                                                 ByVal nLatitude as Double, _
                                                 ByVal nLongitude as Double, _
                                                 ByVal nDepression as Double, _
                                                 ByVal bEarly as BOOLEAN, _
                                                 ByRef bBogus as BOOLEAN) as Double

Dim nMoment        as Double
Dim nPrecision     as Double
Dim nLoop          as BOOLEAN

    nLoop = True
    nPrecision = cmAngle(0,0,30)

    While nLoop = True

        nMoment = cmApproxMomentOfDepression(nApprox,nLatitude,nLongitude,nDepression,bEarly,bBogus)

        Select Case bBogus

        Case False

            If Abs(nApprox - nMoment) < nPrecision Then

                nLoop = False

            Else

                nApprox = nMoment

            End If

        Case Else

            nLoop = False

        End Select
    Wend

    Function = nMoment

End Function
' ========================================================================================
' Approximation for Moment when Sun is at nDepression angle
' ========================================================================================
Private Function cCalendar.cmApproxMomentOfDepression (ByVal nMoment as Double, _
                                                       ByVal nLatitude as Double, _
                                                       ByVal nLongitude as Double, _
                                                       ByVal nDepression as Double, _
                                                       ByVal bEarly as BOOLEAN, _
                                                       ByRef bBogus as BOOLEAN) as Double

' Since the nDepression angle can occur in both in the east (sunrise)
' and west (sunset), bEarly=TRUE for the east and FALSE for the west

' If the nDepression angle doesn't occur, bBogus is set to TRUE

Dim nValue                 as Double
Dim nAlt                   as Double
Dim nTry                   as Double
Dim nDays                  as Long

    bBogus = False
    nDays = cmFloor(nMoment)
    nTry = cmSineOffset(nMoment,nLatitude,nLongitude,nDepression)

    Select Case nDepression

    Case Is >= 0

        nAlt = IIf(bEarly = True,nDays,nDays + 1)

    Case Else

        nAlt = nDays + .5

    End Select

    nValue = IIf(Abs(nTry) > 1,cmSineOffset(nAlt,nLatitude,nLongitude,nDepression),nTry)

    bBogus = IIf(Abs(nValue) > 1,True,False)

    Function = cmLocalFromApparent(nDays + .5 + IIf(bEarly = True,-1,1) * _
                                  (Frac(.5 + cmArcSinDegrees(nValue) / 360) - .25),nLongitude)

End Function
' ========================================================================================
' Atmosphere General Refraction of Light adjusted for Elevation
' ========================================================================================
Private Function cCalendar.cmSolarRefraction (ByVal nElevation as Double, _
                                              ByVal nLatitude as Double) as Double

Dim nEarthRadius    as Double

    nEarthRadius = cmEarthRadius(nLatitude)

    Function = cCalendarClass.VisibleHorizon _
             + cmArcCoSineDegrees(nEarthRadius / (nEarthRadius + IIf(nElevation > 0,nElevation,0))) _
             + cmAngle(0,0,19) * IIf(nElevation > 0,Sqr(nElevation),0)
             
End Function
' ========================================================================================
' Angle between where the sun is at (nMoment) and where we want it to be (nDeclination)
' ========================================================================================
Private Function cCalendar.cmSineOffset (ByVal nMoment as Double, _
                                         ByVal nLatitude as Double, _
                                         ByVal nLongitude as Double, _
                                         ByVal nDepression as Double) as Double

' nMoment is expressed in local time

Dim nUniversal         as Double
Dim nDeclination       as Double

    nUniversal = cmUniversalFromLocal(nMoment,nLongitude)
    nDeclination = cmDeclination(cmObliquity(cmJulianCenturies(nUniversal)),0,cmSolarLongitude(nUniversal))

    Function = cmTangentDegrees(nLatitude) _
             * cmTangentDegrees(nDeclination) _
             + (cmSinDegrees(nDepression) / (cmCoSineDegrees(nDeclination) * cmCoSineDegrees(nLatitude)))

End Function
' ========================================================================================
' Return the Distance in meters of the Moon from Earth
' ========================================================================================
Private Function cCalendar.cmLunarDistance (ByVal nMoment as Double) as Double

Dim nC             as Double   ' Julian Centuries
Dim nElongation    as Double
Dim nSolarAnomaly  as Double
Dim nLunarAnomaly  as Double
Dim nMoonFromNode  as Double
Dim nE             as Double
Dim nCorrection    as Double

   nC = cmJulianCenturies(nMoment)
   nElongation = cmLunarElongation(nC)
   nSolarAnomaly = cmSolarAnomaly(nC)
   nLunarAnomaly = cmLunarAnomaly(nC)
   nMoonFromNode = cmMoonNode(nC)

   nE = 1 - .002516 * nC - .0000074 * nC^2

   nCorrection = cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-20905355,0,0,1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-2955968,2,0,0,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,48888,0,1,0,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,246158,2,0,-2,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-170733,2,0,1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-129620,0,1,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,104755,0,1,1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-34782,4,0,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-21636,4,0,-2,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,30824,2,1,0,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-16675,1,1,0,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-10445,2,0,2,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,14403,2,0,-3,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,6322,1,0,1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,5751,0,1,2,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-4950,2,-2,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,2616,2,1,1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-2117,0,2,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1423,4,0,1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1571,4,-1,0,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,1165,0,2,1,0)

   nCorrection = nCorrection _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-3699111,2,0,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-569925,0,0,2,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-3149,0,0,0,2) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-152138,2,-1,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-204586,2,-1,0,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,108743,1,0,0,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,10321,2,0,0,-2) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,79661,0,0,1,-2) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-23210,0,0,3,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,24208,2,1,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-8379,1,0,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-12831,2,-1,1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-11650,4,0,0,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-7003,0,1,-2,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,10056,2,-1,-2,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-9884,2,-2,0,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,4130,2,0,1,-2) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-3958,4,-1,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,3258,3,0,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1897,4,-1,-2,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,2354,2,2,-1,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1117,0,0,4,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1739,1,0,-2,0) _
               + cmSumDistancePeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-4421,0,0,2,-2)

    Function =  385000560 + nCorrection  ' Distance returned in meters

End Function
' ========================================================================================
' Distance adjustments of the Moon from Earth
' ========================================================================================
Private Function cCalendar.cmSumDistancePeriods (ByRef nE as Double, _
                                                 ByRef nElongation as Double, _
                                                 ByRef nSolarAnomaly as Double, _
                                                 ByRef nLunarAnomaly as Double, _
                                                 ByRef nMoonFromNode as Double, _
                                                 ByVal nV as Double, _
                                                 ByVal nW as Double, _
                                                 ByVal nX as Double, _
                                                 ByVal nY as Double, _
                                                 ByVal nZ as Double) as Double

    Function = nV * nE^Abs(nX) _
             * cmCoSineDegrees((nW * nElongation) + _
                               (nX * nSolarAnomaly) + _
                               (nY * nLunarAnomaly) + _
                               (nZ * nMoonFromNode))

End Function
' ========================================================================================
' Return the Latitude of the Moon
' ========================================================================================
Private Function cCalendar.cmLunarLatitude (ByVal nMoment as Double) as Double

Dim nC                 as Double  ' Julian Centuries
Dim nMeanMoon          as Double
Dim nElongation        as Double
Dim nSolarAnomaly      as Double
Dim nLunarAnomaly      as Double
Dim nMoonFromNode      as Double
Dim nE                 as Double
Dim nVenus             as Double
Dim nFlatEarth         as Double
Dim nExtra             as Double
Dim nCorrection        as Double

    nC = cmJulianCenturies(nMoment)
    nMeanMoon = cmMeanLunarLongitude(nC)
    nElongation = cmLunarElongation(nC)
    nSolarAnomaly = cmSolarAnomaly(nC)
    nLunarAnomaly = cmLunarAnomaly(nC)
    nMoonFromNode = cmMoonNode(nC)
    nE = 1 - .002516 * nC - .0000074 * nC^2
    nVenus = .000175 * (cmSinDegrees(119.75 + (nC * 131.849) + nMoonFromNode) _
            +  cmSinDegrees(119.75 + (nC * 131.849) - nMoonFromNode))
    nFlatEarth = (-.002235 * cmSinDegrees(nMeanMoon)) _
                + (.000127 * cmSinDegrees(nMeanMoon - nLunarAnomaly)) _
                + (-.000115 * cmSinDegrees(nMeanMoon + nLunarAnomaly))
    nExtra = .000382 * cmSinDegrees(313.45 + nC * 481266.484)

    nCorrection = cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,5128122,0,0,0,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,277693,0,0,1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,55413,2,0,-1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,32573,2,0,0,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,9266,2,0,1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,8216,2,-1,0,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,4200,2,0,1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,2463,2,-1,-1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,2065,2,-1,-1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,1828,4,0,-1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1749,0,0,0,3) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1491,1,0,0,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1410,0,1,1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1335,1,0,0,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,1021,4,0,0,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,777,0,0,1,-3) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,607,2,0,0,-3) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,491,2,-1,1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,439,0,0,3,-1)

' Whole series too long handle so break it up

    nCorrection = nCorrection _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,421,2,0,-3,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-351,2,1,0,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,315,2,-1,1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-283,0,0,1,3) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,223,1,1,0,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-220,0,1,-2,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-185,1,0,1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-177,0,1,2,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,166,4,-1,-1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,132,4,0,1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,115,4,-1,0,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,280602,0,0,1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,173237,2,0,0,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,46271,2,0,-1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,17198,0,0,2,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,8822,0,0,2,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,4324,2,0,-2,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-3359,2,1,0,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,2211,2,-1,0,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1870,0,1,-1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1794,0,1,0,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1565,0,1,-1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1475,0,1,1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1344,0,1,0,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,1107,0,0,3,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,833,4,0,-1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,671,4,0,-2,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,596,2,0,2,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-451,2,0,-2,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,422,2,0,2,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-366,2,1,-1,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,331,4,0,0,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,302,2,-2,0,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-229,2,1,1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,223,1,1,0,1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-220,2,1,-1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,181,2,-1,-2,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,176,4,0,-2,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-164,1,0,1,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-119,1,0,-2,-1) _
                 + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,107,2,-2,0,1)

    nCorrection = .000001 * nCorrection

    Function = cmCalcDegrees(nCorrection + nVenus + nFlatEarth + nExtra)

End Function
' ========================================================================================
' Given a date/time moment, this routine will calculate the previous
' lunar event within a longitutinal limitation. Typical lunar events
' are new and full moons.
' ========================================================================================
Private Function cCalendar.cmLunarPhaseAtOrBefore (ByVal nMoment as Double, _
                                                   ByVal nTargetLongitude as Double) as Double

' The basic strategy is to take the moment and number of degrees and
' search for the moment when next the nLongitude of the moon is a multiple of
' the given degrees. The search is a bisection within an interval beginning
' 5 days before our estimate or nMoment whichever is earlier and
' ending long enough past nMoment to insure that the moon passes through
' exactly one multiple of nuTargetLongitude. The process terminates when the
' time is ascertained within one hundred-thousanth of a day (about 0.9 seconds).
' The discontinuity from 360 to 0 degress is taken into account.

Dim nStartMoment       as Double      ' Lower moment of range
Dim nEndMoment         as Double      ' Higher moment of range
Dim nNewMoment         as Double      ' bisection of Lower/Higher range
Dim nNewLongitude      as Double      ' Lunar Longitude of bisection
Dim nSearch            as BOOLEAN

' Calculate upper part of bisection interval

    nEndMoment = nMoment _
               - ((cCalendarClass.MeanSynodicMonth / 360) * cmCalcDegrees(((cmLunarPhase(nMoment) - nTargetLongitude))))

    nStartMoment = nEndMoment - 2
    nEndMoment = IIf(nMoment > nEndMoment + 2,nEndMoment + 2,nMoment)

    nSearch = True

    While nSearch = True

        nNewMoment = nStartMoment + ((nEndMoment - nStartMoment) * .5)
        nNewLongitude = cmCalcDegrees(cmLunarPhase(nNewMoment) - nTargetLongitude)

        Select Case nNewLongitude

        Case Is < 180

            nEndMoment = nNewMoment

        Case Else

            nStartMoment = nNewMoment

        End Select

        nSearch = IIf(nEndMoment - nStartMoment >= .00001,True,False)

    Wend

    Function = nStartMoment + ((nEndMoment - nStartMoment) * .5)

End Function
' ========================================================================================
' Given a date/time moment, this routine will calculate the next
' lunar event within a longitutinal limitation. Typical lunar events
' are new and full moons.
' ========================================================================================
Private Function cCalendar.cmLunarPhaseAtOrAfter (ByVal nMoment as Double, _
                                                  ByVal nTargetLongitude as Double) as Double

' The basic strategy is to take the moment and number of degrees and
' search for the moment when next the nLongitude of the moon is a multiple of
' the given degrees. The search is a bisection within an interval beginning
' 2 days before our estimate or nMoment whichever is earlier and
' ending long enough past nMoment to insure that the moon passes through
' exactly one multiple of nTargetLongitude. The process terminates when the
' time is ascertained within one hundred-thousannh of a day (about 0.9 seconds).

' The discontinuity from 360 to 0 degress is taken into account.

Dim nStartMoment       as Double      ' Lower moment of range
Dim nEndMoment         as Double      ' Higher moment of range
Dim nNewMoment         as Double      ' bisection of Lower/Higher range
Dim nNewLongitude      as Double      ' Lunar Longitude of bisection
Dim nSearch            as Double

' Calculate upper part of bisection interval

    nEndMoment = nMoment + (cCalendarClass.MeanSynodicMonth / 360) _
               * cmCalcDegrees(((nTargetLongitude - cmLunarPhase(nMoment))))

    nStartMoment = IIf(nMoment > nEndMoment - 2,nMoment,nEndMoment - 2)
    nEndMoment = nEndMoment + 2

    nSearch =  True

    Do While nSearch = True

        nNewMoment = nStartMoment + ((nEndMoment - nStartMoment) * .5)

        nNewLongitude = cmCalcDegrees(cmLunarPhase(nNewMoment) - nTargetLongitude)

        Select Case nNewLongitude

        Case Is < 180

            nEndMoment = nNewMoment

        Case Else

            nStartMoment = nNewMoment

        End Select

        nSearch = IIf(nEndMoment - nStartMoment >= .00001,True,False)

    Loop

    Function = nStartMoment + ((nEndMoment - nStartMoment) * .5)

End Function
' ========================================================================================
' Return New Moon following nMoment
' ========================================================================================
Private Function cCalendar.cmNewMoonAfter (ByVal nMoment as Double) as Double

' There are slight differences between the approximations
' used by cmNthNewMoon and cmLunarPhase (which in turn uses
' cmSolarLongitude and cmLunarLongitude) which lead to rare
' occasions (i.e. year 2481) when nMoment is very close to
' the time of a new moon which are addressed.

Dim nNewMoon       as Double
Dim nN0            as Double
Dim nLunarPhase    as Double
Dim nNthMoon       as Long

    nN0 = cmNthNewMoon(0)
    nLunarPhase = cmLunarPhase(nMoment)

' To ensure independence of the phase at the R.D. epoch,
' also subtract from nMoment the moment nN0 of the first
' new moon after R.D. 0.

    nNthMoon = cmFloor(cmRound(((nMoment - nN0) / cCalendarClass.MeanSynodicMonth) - (nLunarPhase / 360)))

    If nLunarPhase < 2 AndAlso cmNthNewMoon(nNthMoon) > nMoment Then

        nNewMoon = cmNthNewMoon(nNthMoon)

    Else

        If nLunarPhase > 358 AndAlso cmNthNewMoon(nNthMoon + 1) <= nMoment Then

            nNewMoon = cmNthNewMoon(nNthMoon + 2)

        Else

            nNewMoon = cmNthNewMoon(nNthMoon + 1)

        End If

    End If

    Function = nNewMoon

End Function
' ========================================================================================
' Return New Moon preceding nMoment
' ========================================================================================
Private Function cCalendar.cmNewMoonBefore (nMoment as Double) as Double

' There are slight differences between the approximations
' used by cmNthNewMoon and cmLunarPhase (which in turn uses
' cmSolarLongitude and cmLunarLongitude) which lead to rare
' occasions (i.e. year 2481) when nMoment is very close to
' the time of a new moon which are addressed.

Dim nNewMoon       as Double
Dim nN0            as Double
Dim nLunarPhase    as Double
Dim nNthMoon       as Long

    nN0 = cmNthNewMoon(0)
    nLunarPhase = cmLunarPhase(nMoment)

' To ensure independence of the phase at the R.D. epoch,
' also subtract from nMoment the moment nN0 of the first
' new moon after R.D. 0.

    nNthMoon = cmRound(((nMoment - nN0) / cCalendarClass.MeanSynodicMonth) - (nLunarPhase / 360))

    If nLunarPhase < 2 AndAlso cmNthNewMoon(nNthMoon) > nMoment Then

        nNewMoon = cmNthNewMoon(nNthMoon - 1)

    Else

        If nLunarPhase > 358 AndAlso cmNthNewMoon(nNthMoon + 1) <= nMoment Then

            nNewMoon = cmNthNewMoon(nNthMoon + 1)

        Else

            nNewMoon = cmNthNewMoon(nNthMoon)

        End If

    End If

    Function = nNewMoon

End Function
' ========================================================================================
' Lunar Phase
' ========================================================================================
Private Function cCalendar.cmLunarPhase (ByVal nMoment as Double) as Double

' Includes check if the phase obtained by the difference between
' lunar and solar longitudes conflicts with the time of the new
' moon as calculated by the more precise cmNthNewMoon function.

' If it does, then an approximation based on cmNthNewMoon is
' preferred.

Dim nLongitudeDifference       as Double
Dim nNthNewMoon                as Double
Dim nMeanSynodic               as Long
Dim nPreferred                 as Double

    nLongitudeDifference = cmCalcDegrees(cmLunarLongitude(nMoment) - cmSolarLongitude(nMoment))
    nNthNewMoon = cmNthNewMoon(0)
    nMeanSynodic = cmFloor(cmRound((nMoment - nNthNewMoon) / cCalendarClass.MeanSynodicMonth))
    nPreferred = Frac((nMoment - cmNthNewMoon(nMeanSynodic)) / cCalendarClass.MeanSynodicMonth) * 360

    Function = IIf(Abs(nLongitudeDifference - nPreferred) > 180,nPreferred,nLongitudeDifference)

End Function
' ========================================================================================
' Return the Longitude of the Moon
' ========================================================================================
Private Function cCalendar.cmLunarLongitude (nMoment as Double) as Double

Dim nC                 as Double  'Julian Centuries
Dim nMeanMoon          as Double
Dim nElongation        as Double
Dim nSolarAnomaly      as Double
Dim nLunarAnomaly      as Double
Dim nMoonFromNode      as Double
Dim nE                 as Double
Dim nVenus             as Double
Dim nJupiter           as Double
Dim nFlatEarth         as Double
Dim nCorrection        as Double

    nC = cmJulianCenturies(nMoment)
    nMeanMoon = cmMeanLunarLongitude(nC)
    nElongation = cmLunarElongation(nC)
    nSolarAnomaly = cmSolarAnomaly(nC)
    nLunarAnomaly = cmLunarAnomaly(nC)
    nMoonFromNode = cmMoonNode(nC)
    nE = 1 - .002516 * nC - .0000074 * nC^2
    nVenus = .003958  * cmSinDegrees(119.75 + (nC * 131.849))
    nJupiter = .000318 * cmSinDegrees(53.09 + (nC * 479264.29))
    nFlatEarth = .001962 * cmSinDegrees(nMeanMoon - nMoonFromNode)

    nCorrection = cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,6288774,0,0,1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,658314,2,0,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-185116,0,1,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,58793,2,0,-2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,53322,2,0,1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-40923,0,1,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-30383,0,1,1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-12528,0,0,1,2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,10675,4,0,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,8548,4,0,-2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-6766,2,1,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,4987,1,1,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,3994,2,0,2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,3665,2,0,-3,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-2602,2,0,-1,2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-2348,1,0,1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-2120,0,1,2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,2048,2,-2,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1595,2,0,0,2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1110,0,0,2,2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-810,2,1,1,0)

' Whole series too long to handle so break it up

    nCorrection = nCorrection _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-713,0,2,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,691,2,1,-2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,549,4,0,1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,520,4,-1,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-399,2,1,0,-2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,351,1,1,1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,330,4,0,-3,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-323,0,2,1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,294,2,0,3,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,1274027,2,0,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,213618,0,0,2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-114332,0,0,0,2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,57066,2,-1,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,45758,2,-1,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-34720,1,0,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,15327,2,0,0,-2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,10980,0,0,1,-2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,10034,0,0,3,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-7888,2,1,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-5163,1,0,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,4036,2,-1,1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,3861,4,0,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-2689,0,1,-2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,2390,2,-1,-2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,2236,2,-2,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-2069,0,2,0,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-1773,2,0,1,-2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,1215,4,-1,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-892,3,0,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,759,4,-1,-2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-700,2,2,-1,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,596,2,-1,0,-2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,537,0,0,4,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-487,1,0,-2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-381,0,0,2,-2) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,-340,3,0,-2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,327,2,-1,2,0) _
                + cmSumLunarPeriods(nE,nElongation,nSolarAnomaly,nLunarAnomaly,nMoonFromNode,299,1,1,-1,0)

    nCorrection = .000001 * nCorrection

    Function = cmCalcDegrees(nMeanMoon + nCorrection + nVenus + nJupiter + nFlatEarth + cmNutation(nC))

End Function
' ========================================================================================
' Adjustments for Longitude of the Moon
' ========================================================================================
Private Function cCalendar.cmSumLunarPeriods (ByRef nE as Double, _
                                              ByRef nElongation as Double, _
                                              ByRef nSolarAnomaly as Double, _
                                              ByRef nLunarAnomaly as Double, _
                                              ByRef nMoonFromNode as Double, _
                                              ByVal nV as Double, _
                                              ByVal nW as Double, _
                                              ByVal nX as Double, _
                                              ByVal nY as Double, _
                                              ByVal nZ as Double) as Double

    Function = nV * nE^Abs(nX) _
             * cmSinDegrees((nW * nElongation) + _
                            (nX * nSolarAnomaly) + _
                            (nY * nLunarAnomaly) + _
                            (nZ * nMoonFromNode))

End Function
' ========================================================================================
' Mean Lunar Longitude
' ========================================================================================
Private Function cCalendar.cmMeanLunarLongitude (ByRef nC as Double) as Double

    Function = cmCalcDegrees(218.3164477 + 481267.88123421 * nC - _
                             .0015786 * nC^2 + (nC^3 / 538841) - _
                             (nC^4 / 65194000))

End Function
' ========================================================================================
' Lunar Elongation
' ========================================================================================
Private Function cCalendar.cmLunarElongation (ByRef nC as Double) as Double

    Function = cmCalcDegrees(297.8501921 + 445267.1114034 * nC - _
                             .0018819 * nC^2 + (nC^3 / 545868) - _
                             (nC^4 / 113065000))

End Function
' ========================================================================================
' Solar Anomaly
' ========================================================================================
Private Function cCalendar.cmSolarAnomaly (ByRef nC as Double) as Double

    Function = cmCalcDegrees(357.5291092 + 35999.0502909 * nC - _
                             .0001536 * nC^2 + (nC^3 / 24490000))

End Function
' ========================================================================================
' Lunar Anomaly
' ========================================================================================
Private Function cCalendar.cmLunarAnomaly (ByRef nC as Double) as Double

    Function = cmCalcDegrees(134.9633964 + 477198.8675055 * nC + _
                             .0087414 * nC^2 + (nC^3 / 69699) - _
                             (nC^4 / 14712000))

End Function
' ========================================================================================
' Moon Node
' ========================================================================================
Private Function cCalendar.cmMoonNode (ByRef nC as Double) as Double

    Function = cmCalcDegrees(93.2720950 + 483202.0175233 * nC - _
                             .0036539 * nC^2 - (nC^3 / 3526000) + _
                             (nC^4 / 863310000))

End Function
' ========================================================================================
' Moment (at Greenwich) of nth new moon after (or before if nNthMoon is negative)
' the new moon of January 11, 1.
' ========================================================================================
Private Function cCalendar.cmNthNewMoon (nNthMoon as Long) as Double

Dim nC                     as Double
Dim nC2                    as Double
Dim nC3                    as Double
Dim nC4                    as Double
Dim nApprox                as Double
Dim nE                     as Double
Dim nK                     as Double
Dim nSolarAnomaly          as Double
Dim nLunarAnomaly          as Double
Dim nMoonArgument          as Double
Dim nOmega                 as Double
Dim nCorrection            as Double
Dim nExtra                 as Double
Dim nAdditional            as Double

    nK = nNthMoon - 24724
    nC = nK / 1236.85
    nC2 = nC^2
    nC3 = nC^3
    nC4 = nC^4

    nApprox = 730125.59766 _
             + cCalendarClass.MeanSynodicMonth * 1236.85 * nC _
             + .0001337 * nC2 _
             - .000000150 * nC3 _
             + .00000000073 * nC4

    nE = 1.0 - .002516 * nC - .0000074 * nC2

    nSolarAnomaly = 2.5534 + 29.10535670 * 1236.85 * nC _
                   - .0000014 * nC2 - .00000011 * nC3

    nLunarAnomaly = 201.5643 + (385.81693528 * 1236.85) * nC _
                   + .0107582 * nC2 + .00001238 * nC3 _
                   - .000000058 * nC4

    nMoonArgument = 160.7108 + (390.67050284 * 1236.85) * nC _
                   - .0016118 * nC2 - .00000227 * nC3 _
                   + .000000011 * nC4

    nOmega = 124.7746 + (-1.56375588 * 1236.85) * nC _
            + .0020672 * nC2 + .00000215 * nC3

    nCorrection = -.00017 * cmSinDegrees(nOmega) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.40720,0,0,1,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.01608,0,0,2,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00739,1,-1,1,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00208,2,2,0,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.00057,0,0,1,2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.00042,0,0,3,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00038,1,1,0,-2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.00007,0,2,1,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00004,0,3,0,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00003,0,0,2,2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00003,0,-1,1,2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.00002,0,1,3,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.17241,1,1,0,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.01039,0,0,0,2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.00514,1,1,1,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.00111,0,0,1,-2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00056,1,1,2,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00042,1,1,0,2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.00024,1,-1,2,0) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00004,0,0,2,-2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00003,0,1,1,-2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.00003,0,1,1,2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,-.00002,0,-1,1,-2) _
                 + cmCorrectionAdjustments(nE,nSolarAnomaly,nLunarAnomaly,nMoonArgument,.00002,0,0,4,0)

    nExtra = .000325 * cmSinDegrees(299.77 + 132.8475848 * nC - .009173 * nC2)

    nAdditional = cmAdditionalAdjustments(nK,251.88,.016321,.000165) _
                + cmAdditionalAdjustments(nK,349.42,36.412478,.000126) _
                + cmAdditionalAdjustments(nK,141.74,53.303771,.000062) _
                + cmAdditionalAdjustments(nK,154.84,7.30686,.000056) _
                + cmAdditionalAdjustments(nK,207.19,.121824,.000042) _
                + cmAdditionalAdjustments(nK,161.72,24.198154,.000037) _
                + cmAdditionalAdjustments(nK,331.55,3.592518,.000023) _
                + cmAdditionalAdjustments(nK,251.83,26.641886,.000164) _
                + cmAdditionalAdjustments(nK,84.66,18.206239,.00011) _
                + cmAdditionalAdjustments(nK,207.14,2.453732,.00006) _
                + cmAdditionalAdjustments(nK,34.52,27.261239,.000047) _
                + cmAdditionalAdjustments(nK,291.34,1.844379,.000040) _
                + cmAdditionalAdjustments(nK,239.56,25.513099,.000035)

    Function = cmUniversalFromDynamical(nApprox + nCorrection + nExtra + nAdditional)

End Function
' ========================================================================================
' nth New Moon Adjustments
' ========================================================================================
Private Function cCalendar.cmCorrectionAdjustments (ByRef dtE as Double, _
                                                    ByRef dtSolarAnomaly as Double, _
                                                    ByRef dtLunarAnomaly as Double, _
                                                    ByRef dtMoonArgument as Double, _
                                                    ByVal dtV as Double, _
                                                    ByVal dtW as Double, _
                                                    ByVal dtX as Double, _
                                                    ByVal dtY as Double, _
                                                    ByVal dtZ as Double) as Double

    Function = dtV * dtE^dtW _
             * cmSinDegrees((dtX * dtSolarAnomaly) + (dtY * dtLunarAnomaly) + (dtZ * dtMoonArgument))

End Function
' ========================================================================================
' nth New Moon Additional Adjustments
' ========================================================================================
Private Function cCalendar.cmAdditionalAdjustments(ByRef stK as Double, _
                                                   ByVal stI as Double, _
                                                   ByVal stJ as Double, _
                                                   ByVal stL as Double) as Double

    Function = stL * cmSinDegrees(stI + stJ * stK)

End Function
' ========================================================================================
' Estimation of when solar longitude reaches nLongitude
' ========================================================================================
Private Function cCalendar.cmEstimatePriorSolarLongitude(ByVal nMoment as Double, _
                                                         ByVal nLongitude as Double) as Double

Dim nEstimate          as Double
Dim nRate              as Double
Dim nError             as Double

    nRate = cmMeanTropicalYear(cmJulianCenturies(nMoment)) / 360

    nEstimate = nMoment - nRate * cmCalcDegrees((cmSolarLongitude(nMoment) - nLongitude))
    nError = cmCalcDegrees((cmSolarLongitude(nEstimate) - nLongitude + 180)) - 180
    nEstimate = nEstimate - nRate * nError

    Function = IIf(nMoment < nEstimate,nMoment,nEstimate)

End Function
' ========================================================================================
' Get one of the seasonal equinoxes as a UTC Moment Type
' ========================================================================================
Private Function cCalendar.cmSeasonalEquinox (ByVal nYear as Long, _
                                              ByVal nEquinox as Long) as Double

Dim nMoment            as Double
Dim nTargetMonth       as Long

    Select Case nEquinox

    Case cCalendarClass.SPRING

        nTargetMonth = cCalendarClass.MARCH

    Case cCalendarClass.SUMMER

        nTargetMonth = cCalendarClass.JUNE

    Case cCalendarClass.AUTUMN

        nTargetMonth = cCalendarClass.SEPTEMBER

    Case Else

        nTargetMonth = cCalendarClass.DECEMBER

    End Select

    Function = cmSolarLongitudeAfter(cmDaysFromGregorian(nTargetMonth,15,nYear),nEquinox)

End Function
' ========================================================================================
' Solar Longitude after a given moment
' ========================================================================================
Private Function cCalendar.cmSolarLongitudeAfter (ByVal nMoment as Double, _
                                                  ByVal nTargetLongitude as Double) as Double

' Given date/time moment, this routine will calculate the next
' solar event within a longitutinal limitation. Typical solar events
' are the equinoxes and solstices.

' Vernal (spring) equinox 0 solar longitude on or about March 21
' Summer solstice 90 solar longitude on or about June 21
' Autumnal (fall) equinox 180 solar longitude on or about September 23
' Winter solstice 270 solar longitude on or about December 22.

' The basic strategy is to take the moment and number of degrees and
' search for the moment when next the longitude of the sun is a multiple of
' the given degrees. The search is a bisection within an interval beginning
' 5 days before our estimate or nMoment whichever is earlier and
' ending long enough past nMoment to insure that the sun passes through
' exactly one multiple of nTargetLongitude. The process terminates when the
' time is ascertained within one hundred-thousandth of a day (about 0.9 seconds).
' The discontinuity from 360 to 0 degress is taken into account.

Dim  nStartMoment      as Double      ' Lower moment of range
Dim  nEndMoment        as Double      ' Higher moment of range
Dim  nNewMoment        as Double      ' bisection of Lower/Higher range
Dim  nNewLongitude     as Double      ' Solar Longitude of bisection
Dim  nSearch           as BOOLEAN


' Calculate upper part of bisection interval

    nEndMoment = nMoment _
                + (cmMeanTropicalYear(cmJulianCenturies(nMoment)) / 360) _
                * cmCalcDegrees(((nTargetLongitude - cmSolarLongitude(nMoment))))

    nStartMoment = IIf(nMoment > nEndMoment - 5,nEndMoment - 5,nMoment)
    nEndMoment = nEndMoment + 5
    nSearch =  True

    Do While nSearch = True

        nNewMoment = nStartMoment + ((nEndMoment - nStartMoment) * .5)
        nNewLongitude = cmCalcDegrees(cmSolarLongitude(nNewMoment) - nTargetLongitude)

        Select Case nNewLongitude

        Case Is < 180

            nEndMoment = nNewMoment

        Case Else

            nStartMoment = nNewMoment

        End Select

        nSearch = IIf(nEndMoment - nStartMoment >= .00001,True,False)

    Loop

    Function = nStartMoment + ((nEndMoment - nStartMoment) * .5)

End Function
' ========================================================================================
' Solar Longitude
' ========================================================================================
Private Function cCalendar.cmSolarLongitude (ByVal dtMoment as Double) as Double

Dim dtC                 as Double      ' Julian Centuries
Dim dtLongitude         as Double

    dtC = cmJulianCenturies(dtMoment)

    dtLongitude = 282.7771834 + 36000.76953744 * dtC _
                + (.000005729577951308232 _
                * (cmSumSolarLongitudePeriods(dtC,403406,270.54861,.9287892) _
                + cmSumSolarLongitudePeriods(dtC,195207,340.19128,35999.1376958) _
                + cmSumSolarLongitudePeriods(dtC,119433,63.91854,35999.4089666) _
                + cmSumSolarLongitudePeriods(dtC,112392,331.2622,35998.7287385) _
                + cmSumSolarLongitudePeriods(dtC,3891,317.843,71998.20261) _
                + cmSumSolarLongitudePeriods(dtC,2819,86.631,71998.4403) _
                + cmSumSolarLongitudePeriods(dtC,1721,240.052,36000.35726) _
                + cmSumSolarLongitudePeriods(dtC,660,310.26,71997.4812) _
                + cmSumSolarLongitudePeriods(dtC,350,247.23,32964.4678) _
                + cmSumSolarLongitudePeriods(dtC,334,260.87,-19.4410) _
                + cmSumSolarLongitudePeriods(dtC,314,297.82,445267.1117) _
                + cmSumSolarLongitudePeriods(dtC,268,343.14,45036.884) _
                + cmSumSolarLongitudePeriods(dtC,242,166.79,3.1008) _
                + cmSumSolarLongitudePeriods(dtC,234,81.53,22518.4434) _
                + cmSumSolarLongitudePeriods(dtC,158,3.5,-19.9739) _
                + cmSumSolarLongitudePeriods(dtC,132,132.75,65928.9345) _
                + cmSumSolarLongitudePeriods(dtC,129,182.95,9038.0293) _
                + cmSumSolarLongitudePeriods(dtC,114,162.03,3034.7684) _
                + cmSumSolarLongitudePeriods(dtC,99,29.8,33718.148) _
                + cmSumSolarLongitudePeriods(dtC,93,266.4,3034.448) _
                + cmSumSolarLongitudePeriods(dtC,86,249.2,-2280.773) _
                + cmSumSolarLongitudePeriods(dtC,78,157.6,29929.992) _
                + cmSumSolarLongitudePeriods(dtC,72,257.8,31556.493) _
                + cmSumSolarLongitudePeriods(dtC,68,185.1,149.588) _
                + cmSumSolarLongitudePeriods(dtC,64,69.9,9037.75) _
                + cmSumSolarLongitudePeriods(dtC,46,8,107997.405) _
                + cmSumSolarLongitudePeriods(dtC,38,197.1,-4444.176) _
                + cmSumSolarLongitudePeriods(dtC,37,250.4,151.771) _
                + cmSumSolarLongitudePeriods(dtC,32,65.3,67555.316) _
                + cmSumSolarLongitudePeriods(dtC,29,162.7,31556.08) _
                + cmSumSolarLongitudePeriods(dtC,28,341.5,-4561.54) _
                + cmSumSolarLongitudePeriods(dtC,27,98.5,1221.655) _
                + cmSumSolarLongitudePeriods(dtC,27,291.6,107996.706) _
                + cmSumSolarLongitudePeriods(dtC,25,146.7,62894.167) _
                + cmSumSolarLongitudePeriods(dtC,24,110,31437.369) _
                + cmSumSolarLongitudePeriods(dtC,21,342.6,-31931.757) _
                + cmSumSolarLongitudePeriods(dtC,21,5.2,14578.298) _
                + cmSumSolarLongitudePeriods(dtC,20,230.9,34777.243) _
                + cmSumSolarLongitudePeriods(dtC,18,256.1,1221.999) _
                + cmSumSolarLongitudePeriods(dtC,17,45.3,62894.511) _
                + cmSumSolarLongitudePeriods(dtC,14,242.9,-4442.039) _
                + cmSumSolarLongitudePeriods(dtC,13,151.8,119.066) _
                + cmSumSolarLongitudePeriods(dtC,13,115.2,107997.909) _
                + cmSumSolarLongitudePeriods(dtC,13,285.3,16859.071) _
                + cmSumSolarLongitudePeriods(dtC,12,53.3,-4.578) _
                + cmSumSolarLongitudePeriods(dtC,10,205.7,-39.127) _
                + cmSumSolarLongitudePeriods(dtC,10,126.6,26895.292) _
                + cmSumSolarLongitudePeriods(dtC,10,85.9,12297.536) _
                + cmSumSolarLongitudePeriods(dtC,10,146.1,90073.778)))

    Function = cmCalcDegrees(dtLongitude + cmAberration(dtC) + cmNutation(dtC))

End Function
' ========================================================================================
' Support for adjustment of solar longitude calculation
' ========================================================================================
Private Function cCalendar.cmSumSolarLongitudePeriods (ByRef dtC as Double, _
                                                       ByVal dwX as Double, _
                                                       ByVal dtY as Double, _
                                                       ByVal dtZ as Double) as Double

    Function = dwX * cmSinDegrees(dtY + (dtZ * dtC))

End Function
' ========================================================================================
' Abberration = Effect of the sun's apparent motion of moving about 20.47 seconds of arc
'               while its light is traveling towards Earth
' ========================================================================================
Private Function cCalendar.cmAberration (ByVal nC as Double) as Double

    Function = (.0000974 * cmCoSineDegrees(177.63 + 35999.01848 * nC)) - .005575

End Function
' ========================================================================================
' Nutation = Wobble of the Earth
' ========================================================================================
Private Function cCalendar.cmNutation (ByVal nC as Double) as Double

Dim nC2                as Double
Dim nA                 as Double
Dim nB                 as Double

    nC2 = nC^2
    nA = 124.90 - 1934.134 * nC + .002063 * nC2
    nB = 201.11 + 72001.5377 * nC + .00057 * nC2

    Function = (-.004778 * cmSinDegrees(nA)) - (.0003667 * cmSinDegrees(nB))

End Function
' ========================================================================================
' Angular distance of a point north or south of the celestial equator
' ========================================================================================
Private Function cCalendar.cmDeclination (ByVal nObliquity as Double, _
                                          ByVal nLatitude as Double, _
                                          ByVal nLongitude as Double) as Double

    Function = cmArcSinDegrees(cmSinDegrees(nLatitude) * _
                   cmCoSineDegrees(nObliquity) + _
                   cmCoSineDegrees(nLatitude) * _
                   cmSinDegrees(nObliquity) * _
                   cmSinDegrees(nLongitude))

End Function
' ========================================================================================
' Angular distance measured eastward along the celestial equator from the vernal equinox 
' ========================================================================================
Private Function cCalendar.cmRightAscension (ByVal nObliquity as Double, _
                                             ByVal nLatitude as Double, _
                                             ByVal nLongitude as Double) as Double

    Function = cmArcTanDegrees((cmSinDegrees(nLongitude) * _
                    cmCoSineDegrees(nObliquity)) - _
                    (cmTangentDegrees(nLatitude) * _
                    cmSinDegrees(nObliquity)),cmCoSineDegrees(nLongitude))
                   
End Function
' ========================================================================================
' Convert Moment Time to Sidereal
' ========================================================================================
Private Function cCalendar.cmSiderealFromMoment (ByVal nMoment as Double) as Double

' Returns Degrees

Dim nC         as Double

    nC = (nMoment - cCalendarClass.J2000) / 36525

    Function = cmCalcDegrees( _
             280.46061837 + 36525 * 360.98564736629 * nC + _
             .000387933 * nC^2 - _
             (nC^3 / 38710000))

End Function
' ========================================================================================
' Mean Interval between Vernal Equinoxes
' ========================================================================================
Private Function cCalendar.cmMeanTropicalYear (ByVal nC as Double) as Double

    Function = 365.2421896698 _
             - (.00000615359 * nC) _
             - (.000000000729 * nC^2) _
             + (.000000000624 * nC^3)
End Function
' ========================================================================================
' Earth Radius at a given Latitude
' ========================================================================================
Private Function cCalendar.cmEarthRadius (ByVal nLatitude as Double) as Double

Dim nLatitudeRadians as Double

' As a safeguard, ensure latitude is in the range of 0-90

     nLatitudeRadians = cmDegreesToRadians(cmMod(Abs(nLatitude),90))^2

' Radius at the equator = 6378137 meters
' Radius at the poles = 6356752.314245 meters

     Function = 6356752.314245 * (1 + nLatitudeRadians)^0.5 _
              / ((6356752.314245^2 / 6378137^2) + nLatitudeRadians)^.5

End Function 
' ========================================================================================
' True midnight of a Solar Day
' ========================================================================================
Private Function cCalendar.cmMidnight (ByVal nMoment as Double, _
                                       ByVal nZone as Double, _
                                       ByVal nLongitude as Double) as Double

    Function = cmStandardFromLocal(cmLocalFromApparent(nMoment,nLongitude), _
                                   nZone,nLongitude)

End Function
' ========================================================================================
' True middle of a Solar Day
' ========================================================================================
Private Function cCalendar.cmMidday (ByVal nMoment as Double, _
                                     ByVal nZone as Double, _
                                     ByVal nLongitude as Double) as Double

    Function = cmStandardFromLocal(cmLocalFromApparent(nMoment + .5,nLongitude), _
                                   nZone,nLongitude)

End Function
' ========================================================================================
' Convert Local time to Apparent
' ========================================================================================
Private Function cCalendar.cmApparentFromLocal (ByVal nMoment as Double, _
                                                ByVal nLongitude as Double) as Double

    Function = nMoment + cmEquationOfTime(cmUniversalFromLocal(nMoment,nLongitude))

End Function
' ========================================================================================
' Convert Apparent time to Local
' ========================================================================================
Private Function cCalendar.cmLocalFromApparent (ByVal nMoment as Double, _
                                                ByVal nLongitude as Double) as Double

    Function = nMoment - cmEquationOfTime(cmUniversalFromLocal(nMoment,nLongitude))

End Function
' ========================================================================================
' Equation of Time
' ========================================================================================
Private Function cCalendar.cmEquationOfTime (ByVal nMoment as Double) as Double

Dim nCenturies             as Double           ' Julian Centuries
Dim nY                     as Double
Dim nLongitude             as Double
Dim nAnomaly               as Double
Dim nEccentricity          as Double
Dim nTimeAdjust            as Double

    nCenturies = cmJulianCenturies(nMoment)

    nlongitude = 280.46645 _
               + 36000.76983 _
               * nCenturies _
               + .0003032 _
               * nCenturies^2

    nAnomaly = cmSolarMeanAnomaly(nCenturies)

    nEccentricity = cmEccentricityEarthOrbit(nCenturies)

    nY = cmTangentDegrees(cmObliquity(nCenturies) / 2)^2

    nTimeAdjust = (1 / (2 * cCalendarClass.PI)) _
                * ((nY * cmSinDegrees(nLongitude * 2)) _
                + (-2 * nEccentricity * cmSinDegrees(nAnomaly)) _
                + (4 * nEccentricity * nY * cmSinDegrees(nAnomaly) _
                * cmCoSineDegrees(nLongitude * 2)) _
                + (-.5 * nY * nY * cmSinDegrees(nLongitude * 4)) _
                + (-1.25 * nEccentricity * nEccentricity _
                * cmSinDegrees(nAnomaly * 2)))

    Function = cmSignum(nTimeAdjust) * IIf(Abs(nTimeAdjust) < .5,Abs(nTimeAdjust),.5)

End Function
' ========================================================================================
' Geometric Mean Anomaly of the Sun
' ========================================================================================
Private Function cCalendar.cmSolarMeanAnomaly (ByVal nC as Double) as Double

    Function = 357.52911 + nC * (35999.05029 - .0001537 * nC)

End Function
' ========================================================================================
' Eccentricity of Earth's orbit
' ========================================================================================
Private Function cCalendar.cmEccentricityEarthOrbit (ByVal nCenturies as Double) as Double

    Function =  .016708634 - nCenturies * (.000042037 + .0000001267 * nCenturies)

End Function
' ========================================================================================
' Obliquity Earth Orbit
' ========================================================================================
Private Function cCalendar.cmObliquity (ByVal nCenturies as Double) as Double

Dim nCenturies100      as Double

    nCenturies100 = nCenturies / 100

    Function = cmAngle(23,26,21.448) _
             - cmAngle(0,0,4680.93) * nCenturies100 _
             - cmAngle(0,0,1.55) * nCenturies100^2 _
             + cmAngle(0,0,1999.25) * nCenturies100^3 _
             - cmAngle(0,0,51.38) * nCenturies100^4 _
             - cmAngle(0,0,249.67) * nCenturies100^5 _
             - cmAngle(0,0,39.05) * nCenturies100^6 _
             + cmAngle(0,0,7.12) * nCenturies100^7 _
             + cmAngle(0,0,27.87) * nCenturies100^8 _
             + cmAngle(0,0,5.79) * nCenturies100^9 _
             + cmAngle(0,0,2.45) * nCenturies100^10 _

End Function
' ========================================================================================
' Return decimal degrees
' ========================================================================================
Private Function cCalendar.cmAngle(ByVal nDegrees as Double, _
                                   ByVal nMinutes as Double, _
                                   ByVal nSeconds as Double) as Double

    Function =  (Abs(nDegrees) + (Abs(nMinutes) / 60) + (Abs(nSeconds) / 3600)) * IIf(nDegrees < 0,-1,1)

End Function
' ========================================================================================
' Sine Degrees converted to Radians
' ========================================================================================
Private Function cCalendar.cmSinDegrees (ByVal nTheta as Double) as Double

    Function = Sin(cmDegreesToRadians(nTheta))

End Function
' ========================================================================================
' Cosine Degrees converted to Radians
' ========================================================================================
Private Function cCalendar.cmCoSineDegrees (ByVal nTheta as Double) as Double

    Function = Cos(cmDegreesToRadians(nTheta))

End Function
' ========================================================================================
' Arc Cosine converted to degrees
' ========================================================================================
Private Function cCalendar.cmArcCoSineDegrees (ByVal nTheta as Double) as Double

    Function = cmRadiansToDegrees(Acos(nTheta))

End Function
' ========================================================================================
' Arc Sine converted to Degrees
' ========================================================================================
Private Function cCalendar.cmArcSinDegrees (ByVal nTheta as Double) as Double

    Function = cmRadiansToDegrees(Asin(nTheta))

End Function
' ========================================================================================
' Tangent degrees converted to Radians
' ========================================================================================
Private Function cCalendar.cmTangentDegrees (ByVal nTheta as Double) as Double

    Function = Tan(cmDegreesToRadians(nTheta))

End Function
' ========================================================================================
' Arc Tangent converted to degrees
' ========================================================================================
Private Function cCalendar.cmArcTanDegrees (ByVal nY as Double, _
                                            ByVal nX as Double) as Double
   
    Function = cmCalcDegrees(cmRadiansToDegrees(ATan2(nY,nX)))

End Function
' ========================================================================================
' Normalize degress within 0-360
' ========================================================================================
Private Function cCalendar.cmCalcDegrees (ByVal nDegrees as Double) as Double

    Function = cmMod(nDegrees,360)

End Function
' ========================================================================================
' Julian Centuries since 2000
' ========================================================================================
Private Function cCalendar.cmJulianCenturies (ByVal nMoment as Double) as Double

    Function = (cmDynamicalFromUniversal(nMoment) - cCalendarClass.J2000) / 36525

End Function
' ========================================================================================
' Convert Universal Time to Dynamical
' ========================================================================================
Private Function cCalendar.cmDynamicalFromUniversal (ByVal nUniversal as Double) as Double

    Function = nUniversal + cmEphemerisCorrection(nUniversal)

End Function
' ========================================================================================
' Convert Dynamical Time to Universal
' ========================================================================================
Private Function cCalendar.cmUniversalFromDynamical (ByVal nDynamical as Double) as Double

    Function = nDynamical - cmEphemerisCorrection(nDynamical)

End Function
' ========================================================================================
' General adjustment for the slowly decreasing rotation of the earth
' ========================================================================================
Private Function cCalendar.cmEphemerisCorrection (ByVal nMoment as Double) as Double

Dim nYear              as Long
Dim nTheta             as Double
Dim nCorrection        as Double

    nYear = cmGregorianYearFromDays(cmFloor(nMoment))

    nTheta = cmGregorianDateDifference(cCalendarClass.JANUARY,1,1900,cCalendarClass.JULY,1,nYear) / 36525.0

    Select Case nYear

    Case 1988 To 2019

        nCorrection = (nYear - 1933) / 86400

    Case 1900 To 1987

        nCorrection = -.00002 _
                    + nTheta * .000297 _
                    + nTheta^2 * .025184 _
                    + nTheta^3 * -.181133 _
                    + nTheta^4 * .553040 _
                    + nTheta^5 * -.861938 _
                    + nTheta^6 * .677066  _
                    + nTheta^7 * -.212591

    Case 1800 To 1899

        nCorrection = -0.000009 _
                    + nTheta * .003844 _
                    + nTheta^2 * .083563 _
                    + nTheta^3 * .865736 _
                    + nTheta^4 * 4.867575 _
                    + nTheta^5 * 15.845535 _
                    + nTheta^6 * 31.332267 _
                    + nTheta^7 * 38.291999 _
                    + nTheta^8 * 28.316289 _
                    + nTheta^9 * 11.636204 _
                    + nTheta^10 * 2.043794

    Case 1700 To 1799

        nYear = nYear - 1700
        nCorrection = (8.118780842 _
                    + nYear * -.005092142 _
                    + nYear^2 * .003336121 _
                    + nYear^3 * -.000026684) / 86400.0

    Case 1620 To 1699

        nYear = nYear - 1600
        nCorrection = (196.58333 _
                    + nYear * -4.0675 _
                    + nYear^2 * .0219167) / 86400.0

    Case Else

        nTheta = cmGregorianDateDifference(cCalendarClass.JANUARY,1,1810, _
                                           cCalendarClass.JANUARY,1,nYear) + .5
        nCorrection = ((nTheta^2 / 41048480) - 15) / 86400.0

    End Select

    Function = nCorrection

End Function
' ========================================================================================
' Convert Standard Time to Local
' ========================================================================================
Private Function cCalendar.cmLocalFromStandard (ByVal nStandard as Double, _
                                                ByVal nZone as Double, _
                                                ByVal nLongitude as Double) as Double

    Function = cmLocalFromUniversal(cmUniversalFromStandard(nStandard,nZone),nLongitude)

End Function
' ========================================================================================
' Convert Local Time to Standard
' ========================================================================================
Private Function cCalendar.cmStandardFromLocal (ByVal nLocal as Double, _
                                                ByVal nZone as Double, _
                                                ByVal nLongitude as Double) as Double

    Function = cmStandardFromUniversal(cmUniversalFromLocal(nLocal,nLongitude),nZone)

End Function
' ========================================================================================
' Convert Standard Time to Universal
' ========================================================================================
Private Function cCalendar.cmUniversalFromStandard (ByVal nStandard as Double, _
                                                    ByVal nZone as Double) as Double

    Function = nStandard - nZone / 24

End Function
' ========================================================================================
' Convert Universal Time to Standard
' ========================================================================================
Private Function cCalendar.cmStandardFromUniversal (ByVal nUniversal as Double, _
                                                    ByVal nZone as Double) as Double

    Function = nUniversal + nZone / 24

End Function
' ========================================================================================
' Convert Universal Time to Local
' ========================================================================================
Private Function cCalendar.cmLocalFromUniversal (ByVal nUniversal as Double, _
                                                 ByVal nLongitude as Double) as Double

    Function = nUniversal + cmZoneFromLongitude(nLongitude)

End Function
' ========================================================================================
' Convert Local Time to Universal
' ========================================================================================
Private Function cCalendar.cmUniversalFromLocal (ByVal nLocal as Double, _
                                                 ByVal nLongitude as Double) as Double

    Function = nLocal - cmZoneFromLongitude(nLongitude)

End Function
' ========================================================================================
' Local mean time zone changes every 15 degrees. Convert to a fraction of a day.
' ========================================================================================
Private Function cCalendar.cmZoneFromLongitude (ByVal nLongitude as Double) as Double

    Function = nLongitude / 360

End Function
' ========================================================================================
' Convert Degrees to Radians
' ========================================================================================
Private Function cCalendar.cmDegreesToRadians(ByVal nDegrees as Double) as Double

    Function = nDegrees * 0.0174532925199433
    
End Function
' ========================================================================================
' Convert Radians to Degrees
' ========================================================================================
Private Function cCalendar.cmRadiansToDegrees(ByVal nRadians as Double) as Double

    Function = nRadians * 57.29577951308232
    
End Function

' ########################################################################################
' Common Support
' ########################################################################################

' ========================================================================================
' Adjust a serial date for daylight savings
' ========================================================================================
Private Function cCalendar.cmDaylightSavings (ByVal nSerial as LongInt, _
                                              ByRef uLocale as LOCATION_LOCALE) as LongInt

    uLocale.bDaylightLightSavingsActive = False
                                              
    If nSerial >= uLocale.DaylightSavingsBegins AndAlso nSerial <  uLocale.DaylightSavingsEnds _
       AndAlso uLocale.bApplyDaylightSavings = True Then

       nSerial = nSerial _
               + (Abs(uLocale.DaylightSavingsMinutes) _
               * cCalendarClass.ONE_MINUTE _
               * cmSignum(uLocale.Zone) _
               * -1)
       uLocale.bDaylightLightSavingsActive = True
     
    End If                                              

    Function = nSerial
    
End Function
' ========================================================================================
' Convert time to serial time which is the number of milliseconds
' since midnight for one day. It's possible to provide a mix of 
' time that exceeds one day.
' ========================================================================================
Private Function cCalendar.cmTimeToSerial(ByVal nHour as Short, _
                                          ByVal nMinute as Short, _
                                          ByVal nSecond as Short, _
                                          ByVal nMillisecond as Short) as Long

    Function = Abs(nHour) * cCalendarClass.ONE_HOUR _
             + Abs(nMinute) * cCalendarClass.ONE_MINUTE _
             + Abs(nSecond) * cCalendarClass.ONE_SECOND _
             + Abs(nMillisecond)
             
End Function
' ========================================================================================
' Extract time from a serial period
' ========================================================================================
Private Sub cCalendar.cmTimeFromSerial(ByVal nTime as Long, _
                                       ByRef nHour as Short, _
                                       ByRef nMinute as Short, _
                                       ByRef nSecond as Short, _
                                       ByRef nMillisecond as Short)

Dim nDays       as Long

' Remove anything that might exceed one day

    nDays = cmFloor(nTime / cCalendarClass.ONE_DAY)
    nTime = nTime - nDays * cCalendarClass.ONE_DAY

' If nTime is < 0 assume it has wrapped back past midnight

    If nTime < 0 Then
    
       nTime = cmFloor(cmMod(nTime,cCalendarClass.ONE_DAY))
       
    End If 

    nHour = cmFloor(nTime / cCalendarClass.ONE_HOUR)
    nTime = nTime - nHour * cCalendarClass.ONE_HOUR
    nMinute = cmFloor(nTime / cCalendarClass.ONE_MINUTE)
    nTime = nTime - nMinute * cCalendarClass.ONE_MINUTE
    nSecond = cmFloor(nTime / cCalendarClass.ONE_SECOND)
    nMillisecond = nTime - nSecond * cCalendarClass.ONE_SECOND
             
End Sub
' ========================================================================================
' Convert Moment to Serial Date
' ========================================================================================
Private Function cCalendar.cmMomentToSerial(ByRef nMoment as Double) as LongInt

Dim nDays   as LongInt
Dim nTime   as Long    

    nDays = cmFloor(nMoment)
    nTime = Abs(Frac(nMoment)) * cCalendarClass.ONE_DAY

    Function = (Abs(nDays) * cCalendarClass.ONE_DAY + nTime) * IIf(nDays < 0,-1,1)    

End Function
' ========================================================================================
' Convert Serial to Moment Date
' ========================================================================================
Private Function cCalendar.cmSerialToMoment(ByVal nSerial as LongInt) as Double

Dim nDays   as Long
Dim nTime   as Long

    cmSerialBreakApart(nSerial,nDays,nTime)
    
    Function = nDays + (nTime / cCalendarClass.ONE_DAY) 

End Function
' ========================================================================================
' Breakapart nSerial representing the number of milliseconds since January 1, 1 at midnight
' ========================================================================================
Private Sub cCalendar.cmSerialBreakApart(ByRef nSerial as LongInt, _
                                         ByRef nDays as Long,      _
                                         ByRef nTime as Long)
                       
' nDays = days since January 1, 0001
' nTime = number of milliseconds for partial day   
                            
   nDays = cmFloor(Abs(nSerial) / cCalendarClass.ONE_DAY) * IIf(nSerial < 0,-1,1)
   nTime = cmMod(Abs(nSerial),cCalendarClass.ONE_DAY)

End Sub

' ========================================================================================
' Variation of x MOD y for Real Numbers adjusted so that the modulus
' of a multiple of the divisor is the divisor itself rather than zero.
' 
' If x MOD y = 0 then result is adjusted to y
' ========================================================================================
Private Function cCalendar.cmAMod(ByVal x as Double, _
                                  ByVal y as Double) as Double

    Function = x - y * (cmCeiling(x / y) - 1)

End Function
' ========================================================================================
' x MOD y for Real Numbers, y<>0
' ========================================================================================
Private Function cCalendar.cmMod(ByVal x as Double, _
                                 ByVal y as Double) as Double

    Function = x - y * cmFloor(x / y)

End Function
' ========================================================================================
' Round x up to nearest integer
' ========================================================================================
Private Function cCalendar.cmRound(ByVal x as Double) as Long

    Function = cmFloor(x + .5)

End Function
' ========================================================================================
' Return smallest integer greater than or equal to x
' ========================================================================================
Private Function cCalendar.cmCeiling(ByVal x as Double) as Long

' Small performance boost

'    Function = cmFloor(x * -1) * -1

    Function = -cmFloor(-x)

End Function
' ========================================================================================
' Return largest integer less than or equal to x
' ========================================================================================
Private Function cCalendar.cmFloor(ByVal x as Double) as Long

    Function = Int(x)

End Function
' ========================================================================================
' Return the sign of nAny
' ========================================================================================
Private Function cCalendar.cmSignum (ByVal nAny as Double) as Long

' Return the sign of nAny

Dim nSign      as Long

    Select Case nAny

    Case Is < 0

        nSign = -1

    Case Is > 0

        nSign = 1

    Case Else

        nSign = 0

    End Select

    Function = nSign

End Function