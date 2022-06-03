'------------------------------------------------------------------------------
'   Author: Knuth Konrad 09.04.2019
'  Changed: -
'------------------------------------------------------------------------------
Imports System.ComponentModel
Imports libBAUtilCore.StringHelper

''' <summary>
''' Enhanced with mostly tourism-related methods/properties <see cref="System.DateTime"/> object.
''' </summary>
''' <remarks>
''' Inherits <see cref="DateTime"/>
''' </remarks>
<DefaultProperty("Date")> <Serializable()> Public Class DateTimeHelper

#Region "Declares"
   ''' <summary>
   ''' Format of IATA date, ddMMM or ddDMMMyy
   ''' </summary>
   Public Enum eIATADateType
      DateLong
      DateShort
   End Enum

   ''' <summary>
   ''' Casing of month abbreviations
   ''' </summary>
   Public Enum eIATADateCasing
      ToLower  ' 01jan
      ToMixed  ' 01Jan
      ToUpper  ' 01JAN
   End Enum

   Private mdtmDate As DateTime
#End Region

#Region "Properties - Private"
#End Region

#Region "Properties - Public"

   ''' <summary>
   ''' Gets a DateTime value that represents the date component of the current <see cref="DateTimeOffset"/> object.
   ''' </summary>
   ''' <returns><see cref="Date"/></returns>
   ''' <remarks>Default property</remarks>
   Public Property [Date] As Date
      Get
         Return mdtmDate
      End Get
      Set(value As DateTime)
         mdtmDate = value
      End Set
   End Property

   ''' <summary>
   ''' Gets a DateTime value that represents the date component of the current <see cref="DateTimeOffset"/> object.
   ''' </summary>
   ''' <returns><see cref="System.DateTime"/></returns>
   Public Property [DateTime] As DateTime
      Get
         Return mdtmDate
      End Get
      Set(value As DateTime)
         mdtmDate = value
      End Set
   End Property

   ''' <summary>
   ''' Gets the day of the month represented by the current <see cref="DateTimeOffset"/> object.
   ''' </summary>
   ''' <returns>
   ''' The day component of the current DateTimeOffset object, expressed as a value between 1 and 31.
   ''' </returns>
   Public ReadOnly Property Day As Integer
      Get
         Return Me.Date.Day
      End Get
   End Property

   ''' <summary>
   ''' Gets the day of the week represented by the current <see cref="DateTimeOffset"/> object.
   ''' </summary>
   ''' <returns>
   ''' One of the enumeration values that indicates the day of the week of the current <see cref="DateTimeOffset"/> object.
   ''' </returns>
   Public ReadOnly Property DayOfWeek As System.DayOfWeek
      Get
         Return Me.Date.DayOfWeek
      End Get
   End Property

   ''' <summary>
   ''' Gets the day of the year represented by the current <see cref="DateTimeOffset"/> object.
   ''' </summary>
   ''' <returns>
   ''' The day of the year of the current <see cref="DateTimeOffset"/> object, expressed as a value between 1 and 366.
   ''' </returns>
   Public ReadOnly Property DayOfYear As Integer
      Get
         Return Me.Date.DayOfYear
      End Get
   End Property

   ''' <summary>
   ''' Gets the hour component of the time represented by the current <see cref="DateTimeOffset"/> object.
   ''' </summary>
   ''' <returns>
   ''' The hour component of the current <see cref="DateTimeOffset"/> object. This property uses a 24-hour clock; the value ranges from 0 to 23.
   ''' </returns>
   Public ReadOnly Property Hour As Integer
      Get
         Return Me.Date.Hour
      End Get
   End Property

   ''' <summary>
   ''' Gets the current date represented by <see cref="Date"/> in IATA long date format, e.g. 01JAN21.
   ''' </summary>
   Public ReadOnly Property IATADateLong As String
      Get
         Return Me.ToDateIATA(eIATADateType.DateLong, eIATADateCasing.ToUpper)
      End Get
   End Property

   ''' <summary>
   ''' Gets the current date represented by <see cref="Date"/> in IATA short date format, e.g. 01JAN.
   ''' </summary>
   Public ReadOnly Property IATADateShort As String
      Get
         Return Me.ToDateIATA(eIATADateType.DateShort, eIATADateCasing.ToUpper)
      End Get
   End Property

   ''' <summary>
   ''' Gets a value that indicates whether the time represented by this instance Is based on local time, Coordinated Universal Time (UTC), Or neither.
   ''' </summary>
   Public ReadOnly Property Kind As DateTimeKind
      Get
         Return Me.Date.Kind
      End Get
   End Property

   ''' <summary>
   ''' Gets the milliseconds component of the date represented by this instance.
   ''' </summary>
   Public ReadOnly Property Millisecond As Integer
      Get
         Return Me.Date.Millisecond
      End Get
   End Property

   ''' <summary>
   ''' Gets the minute component of the date represented by this instance.
   ''' </summary>
   Public ReadOnly Property Minute As Integer
      Get
         Return Me.Date.Minute
      End Get
   End Property

   ''' <summary>
   ''' Gets the month component of the date represented by this instance.
   ''' </summary>
   Public ReadOnly Property Month As Integer
      Get
         Return Me.Date.Month
      End Get
   End Property

   ''' <summary>
   ''' Gets a DateTime object that Is set to the current date And time on this computer, expressed as the local time.
   ''' </summary>
   Public ReadOnly Property Now As DateTime
      Get
         Return DateTime.Now
      End Get
   End Property

   ''' <summary>
   ''' Gets the seconds component of the date represented by this instance.
   ''' </summary>
   Public ReadOnly Property Second As Integer
      Get
         Return Me.Date.Second
      End Get
   End Property

   ''' <summary>
   ''' Gets the number of ticks that represent the date And time of this instance.
   ''' </summary>
   Public ReadOnly Property Ticks As Long
      Get
         Return Me.Date.Ticks
      End Get
   End Property

   ''' <summary>
   ''' Gets the time of day for this instance.
   ''' </summary>
   Public ReadOnly Property TimeOfDay As System.TimeSpan
      Get
         Return Me.Date.TimeOfDay()
      End Get
   End Property

   ''' <summary>
   ''' Gets the current date.
   ''' </summary>
   Public ReadOnly Property Today As DateTime
      Get
         Return DateTime.Today
      End Get
   End Property

   ''' <summary>
   ''' Gets a DateTime object that Is set to the current date And time on this computer, expressed as the Coordinated Universal Time (UTC).
   ''' </summary>
   Public ReadOnly Property UtcNow As DateTime
      Get
         Return DateTime.UtcNow
      End Get
   End Property

   ''' <summary>
   ''' Gets the year component of the date represented by this instance.
   ''' </summary>
   Public ReadOnly Property Year As Integer
      Get
         Return Me.Date.Year
      End Get
   End Property

   ''' <summary>
   ''' Represents the largest possible value of DateTime. This field Is read-only.
   ''' </summary>
   Public ReadOnly Property MaxValue As DateTimeOffset
      Get
         Return DateTime.MaxValue
      End Get
   End Property

   ''' <summary>
   ''' Represents the smallest possible value of DateTime. This field Is read-only.
   ''' </summary>
   Public ReadOnly Property MinValue As DateTimeOffset
      Get
         Return DateTime.MinValue
      End Get
   End Property

#End Region

#Region "Methods - Public"

   ''' <summary>
   ''' Returns a New DateTime that adds the value of the specified TimeSpan to the value of this instance.
   ''' </summary>
   ''' <param name="timeSpan">A positive Or negative time interval.</param>
   ''' <returns>An object whose value Is the sum of the date And time represented by this instance And the time interval represented by value.</returns>
   Public Function Add(ByVal timeSpan As System.TimeSpan) As DateTime
      Return Me.Date.Add(timeSpan)
   End Function

   ''' <summary>
   ''' Returns a New DateTime that adds the specified number of days to the value of this instance.
   ''' </summary>
   ''' <param name="days">A number of whole And fractional days. The value parameter can be negative Or positive.</param>
   ''' <returns>An object whose value Is the sum of the date And time represented by this instance And the number of days represented by value.</returns>
   Public Function AddDays(ByVal days As Double) As DateTime
      Return Me.Date.AddDays(days)
   End Function

   ''' <summary>
   ''' Returns a New DateTime that adds the specified number of hours to the value of this instance.
   ''' </summary>
   ''' <param name="hours">A number of whole And fractional hours. The value parameter can be negative Or positive.</param>
   ''' <returns>Returns a New DateTime that adds the specified number of hours to the value of this instance.</returns>
   Public Function AddHours(ByVal hours As Double) As DateTime
      Return Me.Date.AddHours(hours)
   End Function

   ''' <summary>
   ''' Returns a New DateTime that adds the specified number of milliseconds to the value of this instance.
   ''' </summary>
   ''' <param name="milliseconds">A number of whole And fractional milliseconds. The value parameter can be negative Or positive. Note that this value Is rounded to the nearest integer.</param>
   ''' <returns>An object whose value Is the sum of the date And time represented by this instance And the number of milliseconds represented by value.</returns>
   Public Function AddMilliseconds(ByVal milliseconds As Double) As DateTime
      Return Me.Date.AddMilliseconds(milliseconds)
   End Function

   ''' <summary>
   ''' Returns a New DateTime that adds the specified number of minutes to the value of this instance.
   ''' </summary>
   ''' <param name="minutes">A number of whole And fractional minutes. The value parameter can be negative Or positive.</param>
   ''' <returns>An object whose value Is the sum of the date And time represented by this instance And the number of minutes represented by value.</returns>
   Public Function AddMinutes(ByVal minutes As Double) As DateTime
      Return Me.Date.AddMinutes(minutes)
   End Function

   ''' <summary>
   ''' Returns a New DateTime that adds the specified number of months to the value of this instance.
   ''' </summary>
   ''' <param name="months">A number of months. The months parameter can be negative Or positive.</param>
   ''' <returns>An object whose value Is the sum of the date And time represented by this instance And months.</returns>
   Public Function AddMonths(ByVal months As Integer) As DateTime
      Return Me.Date.AddMonths(months)
   End Function

   ''' <summary>
   ''' Returns a New DateTime that adds the specified number of seconds to the value of this instance.
   ''' </summary>
   ''' <param name="seconds">A number of whole And fractional seconds. The value parameter can be negative Or positive.</param>
   ''' <returns>An object whose value Is the sum of the date And time represented by this instance And the number of seconds represented by value.</returns>
   Public Function AddSeconds(ByVal seconds As Double) As DateTime
      Return Me.Date.AddSeconds(seconds)
   End Function

   ''' <summary>
   ''' Returns a New DateTime that adds the specified number of ticks to the value of this instance.
   ''' </summary>
   ''' <param name="ticks">A number of 100-nanosecond ticks. The value parameter can be positive Or negative.</param>
   ''' <returns>An object whose value Is the sum of the date And time represented by this instance And the time represented by value.</returns>
   Public Function AddTicks(ByVal ticks As Long) As DateTime
      Return Me.Date.AddTicks(ticks)
   End Function

   ''' <summary>
   ''' Returns a New DateTime that adds the specified number of years to the value of this instance.
   ''' </summary>
   ''' <param name="years">A number of years. The value parameter can be negative Or positive.</param>
   ''' <returns>An object whose value Is the sum of the date And time represented by this instance And the number of years represented by value.</returns>
   Public Function AddYears(ByVal years As Integer) As DateTime
      Return Me.Date.AddYears(years)
   End Function

   ''' <summary>
   ''' Compares the value of this instance to a specified DateTime value And returns an integer that indicates whether this instance Is earlier than, the same as, Or later than the specified DateTime value.
   ''' </summary>
   ''' <param name="value">The object to compare to the current instance.</param>
   ''' <returns>The object to compare to the current instance.
   ''' Less than zero: This instance Is earlier than value. 
   ''' Zero: This instance Is the same As value. 
   ''' Greater than zero This instance Is later than value. 
   ''' </returns>
   Public Overloads Function CompareTo(ByVal value As DateTime) As Integer
      Return Me.Date.CompareTo(value)
   End Function

   ''' <summary>
   ''' Compares the value of this instance to a specified object that contains a specified DateTime value, And returns an integer that indicates whether this instance Is earlier than, the same as, Or later than the specified DateTime value.
   ''' </summary>
   ''' <param name="value">A boxed object to compare, Or null.</param>
   ''' <returns>The object to compare to the current instance.
   ''' Less than zero: This instance Is earlier than value. 
   ''' Zero: This instance Is the same As value. 
   ''' Greater than zero This instance Is later than value. 
   ''' </returns>
   Public Overloads Function CompareTo(ByVal value As Object) As Integer
      Return Me.Date.CompareTo(value)
   End Function

   ''' <summary>
   ''' Returns the number of days in the specified month And year.
   ''' </summary>
   ''' <param name="year">The year.</param>
   ''' <param name="month">The month (a number ranging from 1 to 12).</param>
   ''' <returns>The number of days in month for the specified year.
   ''' For example, if month equals 2 for February, the return value Is 28 Or 29 depending upon whether year Is a leap year.
   ''' </returns>
   Public Function DaysInMonth(ByVal year As Integer, ByVal month As Integer) As Integer
      Return DateTime.DaysInMonth(year, month)
   End Function

   ''' <summary>
   ''' Returns a value indicating whether the value of this instance Is equal to the value of the specified DateTime instance.
   ''' </summary>
   ''' <param name="value">The object to compare to this instance.</param>
   ''' <returns>
   ''' <see langword="true"/> if the value parameter equals the value of this instance; otherwise, <see langword="false"/>. 
   ''' </returns>
   Public Shadows Function Equals(ByVal value As DateTime) As Boolean
      Return Me.Date.Equals(value)
   End Function

   ''' <summary>
   ''' Returns a value indicating whether two DateTime objects, Or a DateTime instance And another object Or DateTime, have the same value.
   ''' </summary>
   ''' <param name="t1">The first object to compare.</param>
   ''' <param name="t2">The second object to compare.</param>
   ''' <returns>
   ''' <see langword="true"/> if the two values are equal; otherwise, <see langword="false"/>.
   ''' </returns>

   Public Shadows Function Equals(ByVal t1 As DateTime, ByVal t2 As DateTime) As Boolean
      Return DateTime.Equals(t1, t2)
   End Function

   ''' <summary>
   ''' Deserializes a 64-bit binary value And recreates an original serialized DateTime object.
   ''' </summary>
   ''' <param name="dateData">A 64-bit signed integer that encodes the Kind property in a 2-bit field And the Ticks property in a 62-bit field.</param>
   ''' <returns>An object that Is equivalent to the DateTime object that was serialized by the <see cref="DateTime.ToBinary()"/> method.</returns>
   Public Function FromBinary(ByVal dateData As Long) As DateTime
      Return DateTime.FromBinary(dateData)
   End Function

   ''' <summary>
   ''' Converts the specified Windows file time to an equivalent local time.
   ''' </summary>
   ''' <param name="fileTime">A Windows file time expressed in ticks.</param>
   ''' <returns>An object that represents the local time equivalent of the date And time represented by the fileTime parameter.</returns>
   Public Function FromFileTime(ByVal fileTime As Long) As DateTime
      Return DateTime.FromFileTime(fileTime)
   End Function

   ''' <summary>
   ''' Converts the specified Windows file time to an equivalent UTC time.
   ''' </summary>
   ''' <param name="fileTime">A Windows file time expressed in ticks.</param>
   ''' <returns>An object that represents the UTC time equivalent of the date And time represented by the fileTime parameter.</returns>
   Public Function FromFileTimeUtc(ByVal fileTime As Long) As DateTime
      Return DateTime.FromFileTimeUtc(fileTime)
   End Function

   ''' <summary>
   ''' Returns a <see cref="DateTime"/> equivalent to the specified OLE Automation Date.
   ''' </summary>
   ''' <param name="d">An OLE Automation Date value.</param>
   ''' <returns>An object that represents the same date And time as d.</returns>
   Public Function FromOADate(ByVal d As Double) As DateTime
      Return DateTime.FromOADate(d)
   End Function

   ''' <summary>
   ''' Converts the value of this instance to all the string representations supported by the standard date And time format specifiers.
   ''' </summary>
   ''' <returns>A string array where each element Is the representation of the value of this instance formatted with one of the standard date And time format specifiers.</returns>
   Public Overloads Function GetDateTimeFormats() As String()
      Return Me.Date.GetDateTimeFormats
   End Function

   ''' <summary>
   ''' Converts the value of this instance to all the string representations supported by the specified standard date And time format specifier.
   ''' </summary>
   ''' <param name="format">A standard date And time format string.</param>
   ''' <returns>A string array where each element Is the representation of the value of this instance formatted with one of the standard date And time format specifiers.</returns>
   Public Overloads Function GetDateTimeFormats(ByVal format As Char) As String()
      Return Me.Date.GetDateTimeFormats(format)
   End Function

   ''' <summary>
   ''' Converts the value of this instance to all the string representations supported by the specified standard date And time format specifier.
   ''' </summary>
   ''' <param name="format">A standard date And time format string.</param>
   ''' <param name="provider">An object that supplies culture-specific formatting information about this instance.</param>
   ''' <returns>A string array where each element Is the representation of the value of this instance formatted with one of the standard date And time format specifiers.</returns>
   Public Overloads Function GetDateTimeFormats(ByVal format As Char, ByVal provider As System.IFormatProvider) As String()
      Return Me.Date.GetDateTimeFormats(format, provider)
   End Function

   ''' <summary>
   ''' Converts the value of this instance to all the string representations supported by the specified standard date And time format specifier.
   ''' </summary>
   ''' <param name="provider">An object that supplies culture-specific formatting information about this instance.</param>
   ''' <returns>A string array where each element Is the representation of the value of this instance formatted with one of the standard date And time format specifiers.</returns>
   Public Overloads Function GetDateTimeFormats(ByVal provider As System.IFormatProvider) As String()
      Return Me.Date.GetDateTimeFormats(provider)
   End Function

   ''' <summary>
   ''' Indicates whether this instance of <see cref="System.DateTime"/> Is within the daylight saving time range for the current time zone.
   ''' </summary>
   ''' <returns>
   ''' <see langword="true"/> if the value of the <see cref="DateTime.Kind"/> property Is <see cref="DateTimeKind.Local"/> Or 
   ''' <see cref="DateTimeKind.Unspecified"/> And the value of this instance of <see cref="System.DateTime"/> Is within the 
   ''' daylight saving time range for the local time zone; false if <see cref="DateTime.Kind"/> Is <see cref="DateTimeKind.Utc"/>.
   ''' </returns>
   Public Function IsDaylightSavingTime() As Boolean
      Return Me.Date.IsDaylightSavingTime
   End Function

   ''' <summary>
   ''' Returns an indication whether the specified year Is a leap year.
   ''' </summary>
   ''' <param name="year">A 4-digit year.</param>
   ''' <returns><see langword="true"/> if year Is a leap year; otherwise, <see langword="false"/> .</returns>
   Public Function IsLeapYear(ByVal year As Integer) As Boolean
      Return DateTime.IsLeapYear(year)
   End Function

   ''' <summary>
   ''' Returns the value that results from subtracting the specified time Or duration from the value of this instance
   ''' </summary>
   ''' <param name="value">The date And time value to subtract.</param>
   ''' <returns>A time interval that Is equal to the date And time represented by this instance minus the date And time represented by value.</returns>
   Public Overloads Function Subtract(ByVal value As Date) As System.TimeSpan
      Return Me.Date.Subtract(value)
   End Function

   ''' <summary>
   ''' Returns a New DateTime that subtracts the specified duration from the value of this instance.
   ''' </summary>
   ''' <param name="value">The time interval to subtract.</param>
   ''' <returns>An object that Is equal to the date And time represented by this instance minus the time interval represented by value.</returns>
   Public Overloads Function Subtract(ByVal value As System.TimeSpan) As DateTime
      Return Me.Date.Subtract(value)
   End Function

   ''' <summary>
   ''' Return the last day in a month
   ''' </summary>
   ''' <param name="month">Last day of this month</param>
   ''' <param name="year">Month in this year</param>
   ''' <returns>
   ''' Last day of given <paramref name="month"/> in <paramref name="year"/> as <see cref="System.DateTime" />
   ''' </returns>
   Public Overloads Shared Function GetLastDayInMonth(ByVal month As Int32, ByVal year As Int32) As DateTime

      If month = 12 Then
         month = 1
         year = year + 1
      Else
         month += 1
      End If

      Dim dtm As New DateTime(year, month, 1)

      Dim tsp As New TimeSpan(1, 0, 0, 0)
      Return dtm.Subtract(tsp)

   End Function

   ''' <summary>
   ''' Converts a IATA date string to a <see cref="System.DateTime"/> type
   ''' </summary>
   ''' <param name="iataDate">IATA date string, e.g. 01JAN or 01JAN19</param>
   ''' <returns><see cref="System.DateTime"/></returns>
   ''' <exception cref="ArgumentOutOfRangeException"></exception>
   Public Function ToDate(ByVal iataDate As String) As DateTime

      ' Safe guards
      If iataDate.Length <> 5 AndAlso iataDate.Length <> 7 Then
         Throw New ArgumentOutOfRangeException("iataDate", "Invalid IATA date. Date must be either 5 or 7 characters, e.g. 01JAN or 01JAN19.")
      End If

      Dim sDay As String = String.Empty
      Dim sMonth As String = String.Empty, lMonth As Int32
      Dim sYear As String = String.Empty

      ' Short IATA date
      If iataDate.Length = 5 Then
         sDay = Left(iataDate, 2)
         sMonth = Right(iataDate, 3)
         sYear = DateTime.Now.Year.ToString
      ElseIf iataDate.Length = 7 Then
         sDay = Left(iataDate, 2)
         sMonth = iataDate.Substring(2, 3)
         sYear = Right(iataDate, 2)
      End If

      Select Case sMonth.ToLower
         Case "jan"
            lMonth = 1
         Case "feb"
            lMonth = 2
         Case "mar"
            lMonth = 3
         Case "apr"
            lMonth = 4
         Case "may"
            lMonth = 5
         Case "jun"
            lMonth = 6
         Case "jul"
            lMonth = 7
         Case "aug"
            lMonth = 8
         Case "sep"
            lMonth = 9
         Case "oct"
            lMonth = 10
         Case "nov"
            lMonth = 11
         Case "dec"
            lMonth = 12
         Case Else
            Throw New ArgumentOutOfRangeException("iataDate (month)", "Month must be notated as 3-letter English month abbreviations, e.g. OCT for October")
      End Select

      ' Should handle all other exceptions
      Return New DateTime(Integer.Parse(sYear), lMonth, Integer.Parse(sDay))

   End Function

   ''' <summary>
   ''' Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object.
   ''' </summary>
   ''' <returns>A 64-bit signed integer that encodes the Kind And Ticks properties.</returns>
   Public Function ToBinary() As Long
      Return Me.Date.ToBinary
   End Function

   ''' <summary>
   ''' Converts a <see cref="System.DateTime"/> to a IATA date string
   ''' </summary>
   ''' <returns><see cref="System.DateTime"/>Current <see cref="Date"/> as a string format (ddMMM)</returns>
   Public Overloads Function ToDateIATA() As String

      ' Upper case is default
      Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToUpper

   End Function

   ''' <summary>
   ''' Converts a <see cref="System.DateTime"/> to a IATA date string
   ''' </summary>
   ''' <param name="dateType">
   ''' Format of IATA date <see cref="eIATADateType"/>
   ''' </param>
   ''' <returns><see cref="System.DateTime"/>Current <see cref="Date"/> as a string format (ddMMM)</returns>
   Public Overloads Function ToDateIATA(ByVal dateType As eIATADateType) As String

      ' Upper case is default
      If dateType = eIATADateType.DateLong Then
         Return Me.Date.ToString("ddMMMyy", Globalization.CultureInfo.InvariantCulture).ToUpper
      Else
         Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToUpper
      End If

   End Function

   ''' <summary>
   ''' Converts a <see cref="System.DateTime"/> to a IATA date string
   ''' </summary>
   ''' <param name="nameCasing">
   ''' Casing of IATA date <see cref="eIATADateCasing"/>
   ''' </param>
   ''' <returns><see cref="System.DateTime"/>Current <see cref="Date"/> as a string format (ddMMM)</returns>
   Public Overloads Function ToDateIATA(ByVal nameCasing As eIATADateCasing) As String

      Select Case nameCasing
         Case eIATADateCasing.ToLower
            Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToLower
         Case eIATADateCasing.ToMixed
            Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture)
         Case eIATADateCasing.ToUpper
            Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToUpper
         Case Else
            Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToUpper
      End Select

   End Function

   ''' <summary>
   ''' Converts a <see cref="System.DateTime"/> to a IATA date string
   ''' </summary>
   ''' <param name="dateType">
   ''' Format of IATA date <see cref="eIATADateType"/>
   ''' </param>
   ''' <param name="nameCasing">
   ''' Casing of IATA date <see cref="eIATADateCasing"/>
   ''' </param>
   ''' <returns><see cref="System.DateTime"/>Current <see cref="Date"/> as a string format (ddMMM)</returns>
   Public Overloads Function ToDateIATA(ByVal dateType As eIATADateType, ByVal nameCasing As eIATADateCasing) As String

      Dim sResult As String = Me.ToDateIATA(dateType)

      Select Case nameCasing
         Case eIATADateCasing.ToLower
            Return sResult.ToLower()
         Case eIATADateCasing.ToMixed
            Return sResult
         Case eIATADateCasing.ToUpper
            Return sResult.ToUpper
         Case Else
            Return sResult.ToUpper
      End Select

   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to a Windows file time.
   ''' </summary>
   ''' <returns>The value of the current DateTime object expressed as a Windows file time.</returns>
   Public Function ToFileTime() As Long
      Return Me.Date.ToFileTime
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to a Windows file time.
   ''' </summary>
   ''' <returns>The value of the current DateTime object expressed as a Windows file time.</returns>
   Public Function ToFileTimeUtc() As Long
      Return Me.Date.ToFileTimeUtc
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to local time.
   ''' </summary>
   ''' <returns>
   ''' An object whose Kind property Is Local, And whose value Is the local time equivalent to the value of the current DateTime object, 
   ''' Or MaxValue if the converted value Is too large to be represented by a DateTime object, Or MinValue if the converted value Is 
   ''' too small to be represented as a DateTime object.
   ''' </returns>
   Public Function ToLocalTime() As Date
      Return Me.Date.ToLocalTime
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to its equivalent long date string representation.
   ''' </summary>
   ''' <returns>A string that contains the long date string representation of the current DateTime object.</returns>
   Public Function ToLongDateString() As String
      Return Me.Date.ToLongDateString
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to its equivalent long time string representation.
   ''' </summary>
   ''' <returns>A string that contains the long time string representation of the current DateTime object.</returns>
   Public Function ToLongTimeString() As String
      Return Me.Date.ToLongTimeString
   End Function

   ''' <summary>
   ''' Converts the value of this instance to the equivalent OLE Automation date.
   ''' </summary>
   ''' <returns>A double-precision floating-point number that contains an OLE Automation date equivalent to the value of this instance.</returns>
   Public Function ToOADate() As Double
      Return Me.Date.ToOADate
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to its equivalent short date string representation.
   ''' </summary>
   ''' <returns>A string that contains the short date string representation of the current DateTime object.</returns>
   Public Function ToShortDateString() As String
      Return Me.Date.ToShortDateString
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to its equivalent short time string representation.
   ''' </summary>
   ''' <returns>A string that contains the short time string representation of the current DateTime object.</returns>
   Public Function ToShortTimeString() As String
      Return Me.Date.ToShortTimeString
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to its equivalent string representation using the formatting conventions of the current culture.
   ''' </summary>
   ''' <returns>A string representation of the value of the current DateTime object.</returns>
   Public Overrides Function ToString() As String
      Return Me.Date.ToString
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to its equivalent string representation using the specified format And the formatting conventions of the current culture.
   ''' </summary>
   ''' <param name="format">A standard Or custom date And time format string.</param>
   ''' <returns>A string representation of value of the current DateTime object as specified by format.</returns>
   Public Overloads Function ToString(ByVal format As String) As String
      Return Me.Date.ToString(format)
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to its equivalent string representation.
   ''' </summary>
   ''' <param name="provider">An object that supplies culture-specific formatting information.</param>
   ''' <returns>A string representation of value of the current DateTime object as specified by provider.</returns>
   Public Overloads Function ToString(ByVal provider As System.IFormatProvider) As String
      Return Me.Date.ToString(provider)
   End Function

   ''' <summary>
   ''' Converts the value of the current DateTime object to its equivalent string representation.
   ''' </summary>
   ''' <param name="format">A standard Or custom date And time format string.</param>
   ''' <param name="provider">An object that supplies culture-specific formatting information.</param>
   ''' <returns>A string representation of value of the current DateTime object as specified by format And provider.</returns>
   Public Overloads Function ToString(ByVal format As String, ByVal provider As System.IFormatProvider) As String
      Return Me.Date.ToString(format, provider)
   End Function

   ''' <summary>
   ''' Implements <see cref="Object.GetHashCode()"/>
   ''' </summary>
   ''' <returns><see cref="Object.GetHashCode()"/></returns>
   Public Overrides Function GetHashCode() As Int32

      Dim hashCode As Int32 = -794484751
      hashCode = hashCode * -1521134295 + Me.Date.GetHashCode()
      hashCode = hashCode * -1521134295 + Me.DateTime.GetHashCode()
      Return hashCode

   End Function


#End Region

#Region "Constructor/Dispose"
   Public Sub New()
      Me.Date = New DateTime
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure using the specified DateTime value.
   ''' </summary>
   ''' <param name="newDate">A date and time.</param>
   Public Sub New(ByVal newDate As DateTime)
      Me.Date = newDate
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure using the specified number of ticks.
   ''' </summary>
   ''' <param name="ticks">A 100-nanosecond units.</param>
   Public Sub New(ByVal ticks As Int64)
      Me.Date = New DateTime(ticks)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure using the specified number of ticks 
   ''' and to Coordinated Universal Time (UTC) or local time.
   ''' </summary>
   ''' <param name="ticks">A 100-nanosecond units.</param>
   ''' <param name="kind">One of the enumeration values that indicates whether ticks specifies a local time, 
   ''' Coordinated Universal Time (UTC), or neither</param>
   Public Sub New(ByVal ticks As Int64, ByVal kind As DateTimeKind)
      Me.Date = New DateTime(ticks, kind)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified year, month, and day.
   ''' </summary>
   ''' <param name="year">The year (1 through 9999).</param>
   ''' <param name="month">The month (1 through 12).</param>
   ''' <param name="day">The day (1 through the number of days in <paramref name="month"/>).</param>
   Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32)
      Me.Date = New Date(year, month, day)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified year, month, and day for the specified calendar.
   ''' </summary>
   ''' <param name="year">The year (1 through the number of years in <paramref name="calendar"/>).</param>
   ''' <param name="month">The month (1 through 12 in <paramref name="calendar"/>).</param>
   ''' <param name="day">The day (1 through the number of days in <paramref name="month"/>).</param>
   ''' <param name="calendar">The calendar that is used to interpret <paramref name="year"/>, <paramref name="month"/>, and <paramref name="day"/>.</param>
   Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32, ByVal calendar As Globalization.Calendar)
      Me.Date = New Date(year, month, day, calendar)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified year, month, day, hour, minute, and second.
   ''' </summary>
   ''' <param name="year">The year (1 through 9999).</param>
   ''' <param name="month">The month (1 through 12).</param>
   ''' <param name="day">The day (1 through the number of days in <paramref name="month"/>).</param>
   ''' <param name="hour">The hours (0 through 23).</param>
   ''' <param name="minute">The minutes (0 through 59).</param>
   ''' <param name="second">The seconds (0 through 59).</param>
   Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, ByVal second As Int32)
      Me.Date = New Date(year, month, day, hour, minute, second)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified year, month, day, hour, minute, and second, 
   ''' and Coordinated Universal Time (UTC) or local time.
   ''' </summary>
   ''' <param name="year">The year (1 through 9999).</param>
   ''' <param name="month">The month (1 through 12).</param>
   ''' <param name="day">The day (1 through the number of days in <paramref name="month"/>).</param>
   ''' <param name="hour">The hours (0 through 23).</param>
   ''' <param name="minute">The minutes (0 through 59).</param>
   ''' <param name="second">The seconds (0 through 59).</param>
   ''' <param name="kind">One of the enumeration values that indicates whether <paramref name="year"/>, <paramref name="month"/>, 
   ''' <paramref name="day"/>, <paramref name="hour"/>, <paramref name="minute"/> and <paramref name="second"/> specify a local time, 
   ''' Coordinated Universal Time (UTC), or neither.</param>
   Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, ByVal second As Int32, ByVal kind As DateTimeKind)
      Me.Date = New Date(year, month, day, hour, minute, second, kind)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified year, month, day, hour, minute, and second, 
   ''' for the specified calendar.
   '''  </summary>
   ''' <param name="year">The year (1 through 9999).</param>
   ''' <param name="month">The month (1 through 12).</param>
   ''' <param name="day">The day (1 through the number of days in <paramref name="month"/>).</param>
   ''' <param name="hour">The hours (0 through 23).</param>
   ''' <param name="minute">The minutes (0 through 59).</param>
   ''' <param name="second">The seconds (0 through 59).</param>
   ''' <param name="calendar">The calendar that is used to interpret <paramref name="year"/>, <paramref name="month"/>, and <paramref name="day"/>.</param>
   Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, ByVal second As Int32, ByVal calendar As Globalization.Calendar)
      Me.Date = New Date(year, month, day, hour, minute, second, calendar)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified year, month, day, hour, minute, and second, and millisecond.
   ''' </summary>
   ''' <param name="year">The year (1 through 9999).</param>
   ''' <param name="month">The month (1 through 12).</param>
   ''' <param name="day">The day (1 through the number of days in <paramref name="month"/>).</param>
   ''' <param name="hour">The hours (0 through 23).</param>
   ''' <param name="minute">The minutes (0 through 59).</param>
   ''' <param name="second">The seconds (0 through 59).</param>
   ''' <param name="millisecond">The milliseconds (0 through 999).</param>
   Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, ByVal second As Int32, ByVal millisecond As Int32)
      Me.Date = New Date(year, month, day, hour, minute, second, millisecond)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified year, month, day, hour, minute, and second, 
   ''' and Coordinated Universal Time (UTC) or local time.
   ''' </summary>
   ''' <param name="year">The year (1 through 9999).</param>
   ''' <param name="month">The month (1 through 12).</param>
   ''' <param name="day">The day (1 through the number of days in <paramref name="month"/>).</param>
   ''' <param name="hour">The hours (0 through 23).</param>
   ''' <param name="minute">The minutes (0 through 59).</param>
   ''' <param name="second">The seconds (0 through 59).</param>
   ''' <param name="millisecond">The milliseconds (0 through 999).</param>
   ''' <param name="kind">One of the enumeration values that indicates whether <paramref name="year"/>, <paramref name="month"/>, 
   ''' <paramref name="day"/>, <paramref name="hour"/>, <paramref name="minute"/> and <paramref name="second"/> specify a local time, 
   ''' Coordinated Universal Time (UTC), or neither.</param>
   Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, ByVal second As Int32, ByVal millisecond As Int32,
                        kind As DateTimeKind)
      Me.Date = New Date(year, month, day, hour, minute, second, millisecond, kind)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified year, month, day, hour, minute, and second, 
   ''' for the specified calendar.
   '''  </summary>
   ''' <param name="year">The year (1 through 9999).</param>
   ''' <param name="month">The month (1 through 12).</param>
   ''' <param name="day">The day (1 through the number of days in <paramref name="month"/>).</param>
   ''' <param name="hour">The hours (0 through 23).</param>
   ''' <param name="minute">The minutes (0 through 59).</param>
   ''' <param name="second">The seconds (0 through 59).</param>
   ''' <param name="millisecond">The milliseconds (0 through 999).</param>
   ''' <param name="calendar">The calendar that is used to interpret <paramref name="year"/>, <paramref name="month"/>, and <paramref name="day"/>.</param>
   Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, ByVal second As Int32, ByVal millisecond As Int32,
                        calendar As Globalization.Calendar)
      Me.Date = New Date(year, month, day, hour, minute, second, millisecond, calendar)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified year, month, day, hour, minute, and second, 
   ''' for the specified calendar.
   '''  </summary>
   ''' <param name="year">The year (1 through 9999).</param>
   ''' <param name="month">The month (1 through 12).</param>
   ''' <param name="day">The day (1 through the number of days in <paramref name="month"/>).</param>
   ''' <param name="hour">The hours (0 through 23).</param>
   ''' <param name="minute">The minutes (0 through 59).</param>
   ''' <param name="second">The seconds (0 through 59).</param>
   ''' <param name="millisecond">The milliseconds (0 through 999).</param>
   ''' <param name="calendar">The calendar that is used to interpret <paramref name="year"/>, <paramref name="month"/>, and <paramref name="day"/>.</param>
   ''' <param name="kind">One of the enumeration values that indicates whether <paramref name="year"/>, <paramref name="month"/>, 
   ''' <paramref name="day"/>, <paramref name="hour"/>, <paramref name="minute"/> and <paramref name="second"/> specify a local time, 
   ''' Coordinated Universal Time (UTC), or neither.</param>
   Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, ByVal second As Int32, ByVal millisecond As Int32,
                        ByVal calendar As Globalization.Calendar, ByVal kind As DateTimeKind)
      Me.Date = New Date(year, month, day, hour, minute, second, millisecond, calendar, kind)
   End Sub

   ''' <summary>
   ''' Initializes a new instance of the <see cref="System.DateTime"/> structure to the specified [year], month, day.
   ''' </summary>
   ''' <param name="iataDate">The year (1 through 9999).</param>
   ''' <remarks><paramref name="iataDate"/> must be either in the form of <see cref="IATADateShort"/> or <see cref="IATADateShort"/></remarks>
   Public Sub New(ByVal iataDate As String)
      Me.Date = Me.ToDate(iataDate)
   End Sub

#End Region

End Class
