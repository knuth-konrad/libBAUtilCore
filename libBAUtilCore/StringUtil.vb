Imports System.Globalization

''' <summary>
''' General purpose string handling/formatting helpers
''' </summary>
Public Class StringUtil

#Region "Declares"
   ' Bytes to <unit> - Function Bytes2FormattedString()
   ''' <summary>
   ''' Data storage units
   ''' </summary>
   ''' <seealso cref="Bytes2FormattedString"/>
   Public Enum eSizeUnits As Int64
      B = 1024L
      KB = B * B
      MB = KB * B
      GB = MB * B
      TB = GB * B
   End Enum
#End Region

   ''' <summary>
   ''' Creates a formatted string representing the size in its proper 'spelled out' unit
   ''' (Bytes, KB etc.)
   ''' </summary>
   ''' <param name="uintBytes">Number of bytes to transform</param>
   ''' <returns>
   ''' Spelled out size, e.g. 1030 -> '1KB'
   ''' </returns>
   ''' <remarks>
   ''' Author: dbasnett<br />
   ''' Source: http://www.vbforums.com/showthread.php?634675-RESOLVED-Bytes-to-MB-etc
   ''' </remarks>
   Public Overloads Shared Function Bytes2FormattedString(ByVal uintBytes As UInt64) As String

      Dim dblInUnits As Double
      Dim sUnits As String = String.Empty, szAsStr As String = String.Empty

      If uintBytes < eSizeUnits.B Then
         szAsStr = uintBytes.ToString("n0")
         sUnits = "Bytes"
      ElseIf uintBytes <= eSizeUnits.KB Then
         dblInUnits = uintBytes / eSizeUnits.B
         szAsStr = dblInUnits.ToString("n1")
         sUnits = "KB"
      ElseIf uintBytes <= eSizeUnits.MB Then
         dblInUnits = uintBytes / eSizeUnits.KB
         szAsStr = dblInUnits.ToString("n1")
         sUnits = "MB"
      ElseIf uintBytes <= eSizeUnits.GB Then
         dblInUnits = uintBytes / eSizeUnits.MB
         szAsStr = dblInUnits.ToString("n1")
         sUnits = "GB"
      Else
         dblInUnits = uintBytes / eSizeUnits.GB
         szAsStr = dblInUnits.ToString("n1")
         sUnits = "TB"
      End If

      Return String.Format("{0} {1}", szAsStr, sUnits)

   End Function

   ''' <summary>
   ''' Creates a formatted string representing the size in its proper spelled out unit
   ''' (Bytes, KB etc.)
   ''' </summary>
   ''' <param name="uintBytes">Number of bytes to transform</param>
   ''' <param name="largestUnitOnly"><see langword="true"/>: return only the largest part, e.g. 5.5 GB -> 5GB</param>
   ''' <returns>
   ''' Spelled out size, e.g. 1030 -> '1KB'
   ''' </returns>
   ''' <remarks>
   ''' Author: dbasnett<br />
   ''' Source: http://www.vbforums.com/showthread.php?634675-RESOLVED-Bytes-to-MB-etc
   ''' </remarks>
   Public Overloads Shared Function Bytes2FormattedString(ByVal uintBytes As UInt64,
                                                          Optional ByVal largestUnitOnly As Boolean = True) As String
      Dim uintDivisor As UInt64
      Dim sUnits As String = String.Empty, szAsStr As String = String.Empty

      If largestUnitOnly = True Then
         Return Bytes2FormattedString(uintBytes)
      End If

      Do While uintBytes > 0

         If (uintBytes \ CType(1024 ^ 4, UInt64)) > 0 Then
            ' TB
            uintDivisor = uintBytes \ CType(1024 ^ 4, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 4, UInt64))
            sUnits = "TB "
            ' Debug.Print(String.Format("TB - uintDivisor: {0}, uintBytes: {1}", uintDivisor, uintBytes))
         ElseIf uintBytes \ CType(1024 ^ 3, UInt64) > 0 Then
            ' GB
            uintDivisor = uintBytes \ CType(1024 ^ 3, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 3, UInt64))
            sUnits = "GB "
            ' Debug.Print(String.Format("GB - uintDivisor: {0}, uintBytes: {1}", uintDivisor, uintBytes))
         ElseIf uintBytes \ CType(1024 ^ 2, UInt64) > 0 Then
            ' MB
            uintDivisor = uintBytes \ CType(1024 ^ 2, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 2, UInt64))
            sUnits = "MB "
            ' Debug.Print(String.Format("MB - uintDivisor: {0}, uintBytes: {1}", uintDivisor, uintBytes))
         ElseIf uintBytes \ CType(1024 ^ 1, UInt64) > 0 Then
            ' KB
            uintDivisor = uintBytes \ CType(1024 ^ 1, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 1, UInt64))
            sUnits = "KB "
            ' Debug.Print(String.Format("KB - uintDivisor: {0}, uintBytes: {1}", uintDivisor, uintBytes))
         Else
            ' B
            uintDivisor = uintBytes \ CType(1024 ^ 0, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 0, UInt64))
            sUnits = "B "
            ' Debug.Print(String.Format(" B - uintDivisor: {0}, uintBytes: {1}", uintDivisor, uintBytes))
         End If

         szAsStr &= uintDivisor.ToString("n0") & sUnits

      Loop

      Return szAsStr.TrimEnd

   End Function

   ''' <summary>
   ''' Capitalize the first letter of a string.
   ''' </summary>
   ''' <param name="sText">Source string</param>
   ''' <param name="sCulture">Specific culture string e.g. "en-US"</param>
   ''' <returns>
   ''' <paramref name="sText"/> with the first letter capitalized.
   ''' </returns>
   ''' <remarks>
   ''' Source: https://social.msdn.microsoft.com/Forums/vstudio/en-US/c0872f6d-2975-43e6-872a-d2ba7901ed0e/convert-first-letter-of-string-to-capital?forum=csharpgeneral
   ''' </remarks>
   Public Shared Function MCase(ByVal sText As String, Optional ByVal sCulture As String = "") As String

      Dim ti As TextInfo

      Try
         If sCulture.Length > 0 Then
            ti = New CultureInfo(sCulture, False).TextInfo
            Return ti.ToTitleCase(sText)
         Else
            Return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(sText)
         End If
      Catch
         Return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(sText)
      End Try

   End Function

#Region "Method Chr()"
   ''' <summary>
   ''' Replacement for VB6's Chr() function.
   ''' </summary>
   ''' <param name="ansiValue">ANSI value for which to return a string</param>
   ''' <returns>
   ''' ANSI String-Representation of <paramref name="ansiValue"/>
   ''' </returns>
   ''' <remarks>
   ''' Source: https://stackoverflow.com/questions/36976240/c-sharp-char-from-int-used-as-string-the-real-equivalent-of-vb-chr?lq=1
   ''' </remarks>
   Public Overloads Shared Function Chr(ByVal ansiValue As Int32) As String
      Return Char.ConvertFromUtf32(ansiValue)
   End Function

   ''' <summary>
   ''' Replacement for VB6's Chr() function.
   ''' </summary>
   ''' <param name="ansiValue">ANSI value for which to return a string</param>
   ''' <returns>
   ''' ANSI String-Representation of <paramref name="ansiValue"/>
   ''' </returns>
   Public Overloads Shared Function Chr(ByVal ansiValue As UInt32) As String
      ' Return Char.ConvertFromUtf32(CType(ansiValue, Int32))
      Return System.Convert.ToChar(ansiValue).ToString
   End Function

#End Region

#Region "Method InStr()"
   Public Overloads Shared Function InStr(ByVal start As Int32, ByVal string1 As String, ByVal string2 As String) As Int32

      ' Safe guards
      If start <= 0 Then
         Throw New ArgumentOutOfRangeException("start must be > 0.", "start")
      End If

      If String.IsNullOrEmpty(string1) Then
         Return 0
      End If

      If String.IsNullOrEmpty(string2) Then
         Return start
      End If

      Return string1.IndexOf(string2, start - 1) + 1

   End Function

   Public Overloads Shared Function InStr(ByVal string1 As String, ByVal string2 As String) As Int32

      Return InStr(1, string1, string2)

   End Function
#End Region

   ''' <summary>
   ''' Implements VB6's Left$() functionality.
   ''' </summary>
   ''' <param name="source">Source string</param>
   ''' <param name="leftChars">Number of characters to return</param>
   ''' <returns>
   ''' For leftChars ...<br />
   '''    &gt; source.Length: source<br />
   '''    = 0: Empty string<br />
   '''    &lt; 0: Position from the end of source, e.g. Left("1234567890", -2) -> "12345678"
   ''' </returns>
   ''' <remarks>
   ''' Source: Developed from https://stackoverflow.com/questions/844059/net-equivalent-of-the-old-vb-leftstring-length-function/12481156
   ''' </remarks>
   Public Shared Function Left(ByVal source As String, ByVal leftChars As Integer) As String

      If String.IsNullOrEmpty(source) OrElse leftChars = 0 Then
         Return String.Empty
      ElseIf leftChars > source.Length Then
         Return source
      ElseIf leftChars < 0 Then
         Return source.Substring(0, source.Length + leftChars)
      Else
         Return source.Substring(0, Math.Min(leftChars, CType(source.Length, Integer)))
      End If

   End Function

   ''' <summary>
   ''' Implements VB's/PB's Right$() functionality.
   ''' </summary>
   ''' <param name="source">Source string</param>
   ''' <param name="rightChars">Number of characters to return</param>
   ''' <returns>
   ''' For rightChars ...<br />
   '''    &gt; source.Length: source<br />
   '''    = 0: Empty string<br />
   '''    &lt; 0: Position from the start of source, e.g. Right("1234567890", -2) -&gt; "34567890"
   ''' </returns>
   ''' <remarks>
   ''' Source: Developed from https://stackoverflow.com/questions/844059/net-equivalent-of-the-old-vb-leftstring-length-function/12481156
   ''' </remarks>
   Public Shared Function Right(ByVal source As String, ByVal rightChars As Integer) As String

      If String.IsNullOrEmpty(source) OrElse rightChars = 0 Then
         Return String.Empty
      ElseIf rightChars > source.Length Then
         Return source
      ElseIf rightChars < 0 Then
         Return source.Substring(Math.Abs(rightChars))
      Else
         Return source.Substring(source.Length - rightChars, rightChars)
      End If

   End Function

   ''' <summary>
   ''' Implements VB6's/PB's Mid$() functionality, as .NET's String.SubString() 
   ''' differs in its behavior that it raises an exception if startIndex > source.Length, 
   ''' whereas Mid$() returns an empty string in such a case.
   ''' </summary>
   ''' <param name="source">Source string</param>
   ''' <param name="startIndex">(0-based) start</param>
   ''' <param name="length">Number of chars to return</param>
   ''' <returns>
   ''' For <paramref name="startIndex"/> &gt; <paramref name="source"/>.Length: <see cref="String.Empty"/>
   ''' For <paramref name="length"/> &gt; <paramref name="startIndex"/> + <paramref name="source"/>.Length: all of <paramref name="source"/> from <paramref name="startIndex"/>
   ''' </returns>
   ''' <remarks>
   ''' Source: Developed from https://stackoverflow.com/questions/844059/net-equivalent-of-the-old-vb-leftstring-length-function/12481156
   ''' </remarks>
   Public Shared Function Mid(ByVal source As String, ByVal startIndex As Integer, Optional ByVal length As Integer = 0) As String

      ' Safe guards
      If String.IsNullOrEmpty(source) OrElse (startIndex > source.Length) Then
         Return String.Empty
      End If
      If startIndex < 0 Then
         Throw New ArgumentOutOfRangeException("startIndex")
      End If
      If length < 0 Then
         Throw New ArgumentOutOfRangeException("length")
      End If

      ' Adjust length, if needed
      Try
         If startIndex + length > source.Length OrElse length = 0 Then
            Return source.Substring(startIndex - 1)
         Else
            Return source.Substring(startIndex - 1, length)
         End If
      Catch ex As ArgumentOutOfRangeException
         Return String.Empty
      End Try

   End Function

   ''' <summary>
   ''' Encloses <paramref name="text"/> with double quotation marks (").
   ''' </summary>
   ''' <param name="text">Wrap this string in quotation marks.</param>
   ''' <returns><paramref name="text"/> enclosed in double quotation marks (")</returns>
   Public Shared Function EnQuote(ByVal text As String, Optional ByVal quoteChar As String = "") As String
      If quoteChar.Length < 1 Then
         quoteChar = vbQuote()
      End If
      Return quoteChar & text & quoteChar
   End Function

#Region "Method String()"
   ''' <summary>
   ''' Mimics VB6's String() Function
   ''' </summary>
   ''' <param name="character">Character to use</param>
   ''' <param name="count">Number of characters</param>
   ''' <returns>String of <paramref name="count"/> x <paramref name="character"/></returns>
   Public Overloads Shared Function [String](ByVal character As Char, ByVal count As Int32) As String
      Return New String(character, CType(count, Integer))
   End Function

   ''' <summary>
   ''' Mimics VB6's String() Function
   ''' </summary>
   ''' <param name="character">Character to use</param>
   ''' <param name="count">Number of characters</param>
   ''' <returns>String of <paramref name="count"/> x <paramref name="character"/></returns>
   Public Overloads Shared Function [String](ByVal character As Char, ByVal count As UInt32) As String
      Return New String(character, CType(count, Integer))
   End Function

   ''' <summary>
   ''' Mimics VB6's String() Function
   ''' </summary>
   ''' <param name="character">Character to use</param>
   ''' <param name="count">Number of characters</param>
   ''' <returns>String of <paramref name="count"/> x <paramref name="character"/></returns>
   Public Overloads Shared Function [String](ByVal character As String, ByVal count As Int32) As String
      Return New String(CType(character, Char), CType(count, Integer))
   End Function

   ''' <summary>
   ''' Mimics VB6's String() Function
   ''' </summary>
   ''' <param name="character">Character to use</param>
   ''' <param name="count">Number of characters</param>
   ''' <returns>String of <paramref name="count"/> x <paramref name="character"/></returns>
   Public Overloads Shared Function [String](ByVal character As String, ByVal count As UInt32) As String
      Return New String(CType(character, Char), CType(count, Integer))
   End Function

   ''' <summary>
   ''' Remove any occurrence of <paramref name="removeChars"/> in <paramref name="source"/>
   ''' </summary>
   ''' <param name="source">Source string</param>
   ''' <param name="removeChars">List of strings to remove from <paramref name="source"/></param>
   ''' <returns></returns>
   Public Overloads Shared Function TrimAny(ByVal source As String, ByVal removeChars() As Char) As String

      Dim result As String = source

      result = result.TrimStart(removeChars)
      result = result.TrimEnd(removeChars)

      Return result

   End Function

   ''' <summary>
   ''' Remove any occurrence of <paramref name="removeChars"/> in <paramref name="source"/>
   ''' </summary>
   ''' <param name="source">Source string</param>
   ''' <param name="removeChars">List of strings to remove from <paramref name="source"/></param>
   ''' <returns></returns>
   Public Overloads Shared Function TrimAny(ByVal source As String, ByVal removeChars As String) As String

      Dim result As String = source

      result = result.TrimStart(removeChars.ToCharArray)
      result = result.TrimEnd(removeChars.ToCharArray)

      Return result

   End Function

   ''' <summary>
   ''' Remove any occurrence of <paramref name="removeChars"/> in <paramref name="source"/>
   ''' </summary>
   ''' <param name="source">Source string</param>
   ''' <param name="removeChars">List of strings to remove from <paramref name="source"/></param>
   ''' <returns></returns>
   Public Overloads Shared Function TrimAny(ByVal source As String, ByVal removeChars() As String) As String

      Dim result As String = source

      For Each s As String In removeChars
         result = result.TrimStart(s.ToCharArray)
         result = result.TrimEnd(s.ToCharArray)
      Next

      Return result

   End Function

#End Region

#Region "Method Space()"
   ''' <summary>
   ''' Mimics VB6's Space() function
   ''' </summary>
   ''' <param name="count">Number of space</param>
   ''' <returns>String of <paramref name="count"/> spaces</returns>
   Public Overloads Shared Function Space(ByVal count As UInt32) As String
      Return New String(" "c, CType(count, Integer))
   End Function

   ''' <summary>
   ''' Mimics VB6's Space() function
   ''' </summary>
   ''' <param name="count">Number of space</param>
   ''' <returns>String of <paramref name="count"/> spaces</returns>
   Public Overloads Shared Function Space(ByVal count As Int32) As String
      Return New String(" "c, CType(count, Integer))
   End Function

   ''' <summary>
   ''' Mimics VB6's Space() function
   ''' </summary>
   ''' <param name="count">Number of space</param>
   ''' <returns>String of <paramref name="count"/> spaces</returns>
   Public Overloads Shared Function Space(ByVal count As UInt64) As String
      Return New String(" "c, CType(count, Integer))
   End Function

   ''' <summary>
   ''' Mimics VB6's Space() function
   ''' </summary>
   ''' <param name="count">Number of space</param>
   ''' <returns>String of <paramref name="count"/> spaces</returns>
   Public Overloads Shared Function Space(ByVal count As Int64) As String
      Return New String(" "c, CType(count, Integer))
   End Function
#End Region

#Region "Method Unwrap()"

   ''' <summary>
   ''' Remove paired characters from the beginning and end of a string.
   ''' </summary>
   ''' <param name="source">Input string</param>
   ''' <param name="leftChar">Removes the characters in leftChar from the beginning of <paramref name="source"/>, if there is an exact match.</param>
   ''' <param name="rightChar">Removes the characters in rightChar from the end of <paramref name="source"/>, if there is an exact match.</param>
   ''' <returns><paramref name="source"/> with <paramref name="leftChar"/> and <paramref name="rightChar"/> removed</returns>
   ''' <remarks>Mimics PB's Unwrap() method</remarks>
   Public Overloads Shared Function Unwrap(ByVal source As String, ByVal leftChar As String, ByVal rightChar As String) As String

      Dim s As String = String.Empty

      s = source
      If s.Length > leftChar.Length AndAlso Left(s, leftChar.Length) = leftChar Then
         s = Mid(s, leftChar.Length + 1)
      End If

      If s.Length > rightChar.Length AndAlso Right(s, rightChar.Length) = rightChar Then
         s = Left(s, s.Length - rightChar.Length)
      End If

      Return s

   End Function

   ''' <summary>
   ''' Remove paired characters from the beginning and end of a string.
   ''' </summary>
   ''' <param name="source">Input string</param>
   ''' <param name="leftChar">Removes the characters in leftChar from the beginning of <paramref name="source"/>, if there is an exact match.</param>
   ''' <param name="rightChar">Removes the characters in rightChar from the end of <paramref name="source"/>, if there is an exact match.</param>
   ''' <returns><paramref name="source"/> with <paramref name="leftChar"/> and <paramref name="rightChar"/> removed</returns>
   ''' <remarks>Mimics PB's Unwrap() method</remarks>
   Public Overloads Shared Function Unwrap(ByVal source As String, ByVal leftChar As Char, ByVal rightChar As Char) As String

      Dim s As String = String.Empty

      s = source
      If s.Length > 1 AndAlso Left(s, 1) = leftChar Then
         s = Mid(s, 1)
      End If

      If s.Length > 1 AndAlso Right(s, 1) = rightChar Then
         s = Left(s, s.Length - 1)
      End If

      Return s

   End Function

   ''' <summary>
   ''' Remove paired characters from the beginning and end of a string.
   ''' </summary>
   ''' <param name="source">Input string</param>
   ''' <param name="unwrapChar">Removes the characters in unwarapChar from the beginning of  and the end of <paramref name="source"/>, if there is an exact match.</param>
   ''' <returns><paramref name="source"/> with <paramref name="unwrapChar"/> removed</returns>
   ''' <remarks>Mimics PB's Unwrap() method</remarks>
   Public Overloads Shared Function Unwrap(ByVal source As String, ByVal unwrapChar As String) As String

      Dim s As String = String.Empty

      s = source
      If s.Length > unwrapChar.Length AndAlso Left(s, unwrapChar.Length) = unwrapChar Then
         s = Mid(s, unwrapChar.Length + 1)
      End If

      If s.Length > unwrapChar.Length AndAlso Right(s, unwrapChar.Length) = unwrapChar Then
         s = Left(s, s.Length - unwrapChar.Length)
      End If

      Return s

   End Function

   ''' <summary>
   ''' Remove paired characters from the beginning and end of a string.
   ''' </summary>
   ''' <param name="source">Input string</param>
   ''' <param name="unwrapChar">Removes the characters in unwarapChar from the beginning of  and the end of <paramref name="source"/>, if there is an exact match.</param>
   ''' <returns><paramref name="source"/> with <paramref name="unwrapChar"/> removed</returns>
   ''' <remarks>Mimics PB's Unwrap() method</remarks>
   Public Overloads Shared Function Unwrap(ByVal source As String, ByVal unwrapChar As Char) As String

      Dim s As String = String.Empty

      s = source
      If s.Length > 1 AndAlso Left(s, 1) = unwrapChar Then
         s = Mid(s, 1)
      End If

      If s.Length > 1 AndAlso Right(s, 1) = unwrapChar Then
         s = Left(s, s.Length - 1)
      End If

      Return s

   End Function

#End Region

#Region "Date formatting"
   ''' <summary>
   ''' Create a date string of format YYYYMMDD[[T]HHNNSS].
   ''' </summary>
   ''' <param name="dtmDate">Date/Time to format</param>
   ''' <param name="appendTime"><see langref="true"/> = append time to date</param>
   ''' <param name="dateSeparator">Character to separate date parts</param>
   ''' <param name="dateTimeSeparator">Character to separate date part from time part</param>
   ''' <returns>
   ''' Date/time formatted as string.
   ''' </returns>
   Public Shared Function DateYMD(ByVal dtmDate As DateTime, Optional ByVal appendTime As Boolean = False,
                                  Optional ByVal dateSeparator As String = "", Optional ByVal dateTimeSeparator As String = "T") _
                                  As String

      ' Date part
      Dim sResult As String = dtmDate.Year.ToString("0000") & dateSeparator & dtmDate.Month.ToString("00") & dateSeparator & dtmDate.Day.ToString("00")


      ' Time part
      If appendTime = True Then
         sResult &= dateTimeSeparator & dtmDate.Hour.ToString("00") & dtmDate.Minute.ToString("00") & dtmDate.Second.ToString("00")
      End If

      Return sResult

   End Function
#End Region

#Region "VB6 String constants"
   ' ** Replacements for various handy VB6 string constants

   ''' <summary>
   ''' Mimics VB6's vbNewLine constant.
   ''' </summary>
   ''' <param name="n">Return this number of new lines</param>
   ''' <returns>OS-specific new line character(s)</returns>
   Public Shared Function vbNewLine(Optional ByVal n As Int32 = 1) As String
      Dim sResult As String = String.Empty
      For i As Int32 = 1 To n
         sResult &= Environment.NewLine
      Next
      Return sResult
   End Function

   ''' <summary>
   ''' Mimics VB6's vbNullString constant
   ''' </summary>
   ''' <returns>String.Empty</returns>
   Public Shared Function vbNullString() As String
      Return String.Empty
   End Function

   ''' <summary>
   ''' Constant for a double quotation mark (").
   ''' </summary>
   ''' <param name="n">Number of DQs to return</param>
   ''' <returns><paramref name="n"/> "</returns>
   Public Shared Function vbQuote(Optional ByVal n As Int32 = 1) As String
      Dim sResult As String = String.Empty
      For i As Int32 = 1 To n
         sResult &= System.Convert.ToChar(34).ToString
      Next
      Return sResult
   End Function

   ''' <summary>
   ''' Mimics VB6's vbTab constant.
   ''' </summary>
   ''' <param name="n">Number of tabs to return</param>
   ''' <returns><paramref name="n"/> tabs.</returns>
   Public Shared Function vbTab(Optional ByVal n As Int32 = 1) As String
      Dim sResult As String = String.Empty
      For i As Int32 = 1 To n
         sResult &= System.Convert.ToChar(9).ToString
      Next
      Return sResult
   End Function

   ''' <summary>
   ''' Return a string of what's considered to be 'white space'
   ''' </summary>
   ''' <returns></returns>
   Public Shared Function vbWhiteSpace() As String
      Return vbTab() & vbNewLine() & " "
   End Function
#End Region

End Class
