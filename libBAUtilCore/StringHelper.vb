Imports System.Globalization

''' <summary>
''' General purpose string handling/formatting helpers
''' </summary>
Public Class StringHelper

  ' ToDo: implement ParseCount()

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

#Region "Asc()"

  ''' <summary>
  ''' Return the ASCII value of a character
  ''' </summary>
  ''' <param name="text">Return this character's value.</param>
  ''' <param name="startIndex">Start from the specified index.</param>
  ''' <returns>
  ''' ASCII code, e.g. "A" = 65, If <paramref name="text"/> is <see cref="String.Empty" />
  ''' or <paramref name="startIndex"/> &gt; length of <paramref name="text"/>, returns -1 
  ''' </returns>
  ''' <remarks>
  ''' Mimics VB6's/PB's Asc()
  ''' </remarks>
  Public Overloads Shared Function Asc(ByVal text As String, Optional startIndex As Int32 = 1) As Int32

    ' Safe guards
    If text.Length < 1 OrElse startIndex > text.Length Then
      Return -1
    End If

    Return Convert.ToInt32(System.Text.ASCIIEncoding.ASCII.GetBytes(text.ToCharArray, startIndex - 1, 1)(0))

  End Function

  ''' <summary>
  ''' Return the ASCII value of a character
  ''' </summary>
  ''' <param name="text">return this character's value</param>
  ''' <returns>
  ''' ASCII code, e.g. "A" = 65, If <paramref name="text"/> is <see cref="String.Empty" />, 
  ''' returns -1
  ''' </returns>
  ''' <remarks>
  ''' Mimics VB6's/PB's Asc()
  ''' </remarks>
  Public Overloads Shared Function Asc(ByVal text As Char) As Int32

    Return Convert.ToInt32(System.Text.ASCIIEncoding.ASCII.GetBytes(text))

  End Function

  ''' <summary>
  ''' Return the ASCII value of a character
  ''' </summary>
  ''' <param name="text">return this character's value</param>
  ''' <param name="startIndex">Start from the specified index.</param>
  ''' <returns>
  ''' ASCII code, e.g. "A" = 65, If <paramref name="text"/> is <see cref="String.Empty" />
  ''' or <paramref name="startIndex"/> &gt; length of <paramref name="text"/>, returns -1 
  ''' </returns>
  ''' <remarks>
  ''' Mimics VB6's/PB's Asc()
  ''' </remarks>
  Public Overloads Shared Function Asc(ByVal text As Char(), Optional startIndex As Int32 = 1) As Int32

    ' Safe guards
    If text.Length < 1 OrElse startIndex > text.Length Then
      Return -1
    End If

    Return Convert.ToInt32(System.Text.ASCIIEncoding.ASCII.GetBytes(text, startIndex, 1))

  End Function

#End Region

#Region "Bytes2FormattedString"

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

#Region "Extract()"

  ''' <summary>
  ''' Extract characters from a string up to a character or group of characters.
  ''' </summary>
  ''' <param name="mainString">mainString is the string expression from which to extract.</param>
  ''' <param name="anyMatchChar">If <see langword="true"/>, matchString specifies a list of single characters to be searched for individually, 
  ''' a match on any one of which will cause the extract operation to be performed up to that character.
  ''' </param>
  ''' <param name="matchStr">matchString is the string expression to extract up to. Extract is case-sensitive.</param>
  ''' <returns>Extract returns a sub-string of mainString, starting with its first character (or the character specified by startIndex) and up to (but not including) the first occurrence of matchString. 
  ''' If matchString is not present in mainString, or either string parameter is empty, all of mainString is returned.
  ''' </returns>
  ''' <remarks>Mimics PB's Extract$()</remarks>
  Public Overloads Shared Function Extract(ByVal mainString As String, ByVal anyMatchChar As Boolean, ByVal matchStr As String) As String
    Return Extract(1, mainString, anyMatchChar, matchStr)
  End Function

  ''' <summary>
  ''' Extract characters from a string up to a character or group of characters.
  ''' </summary>
  ''' <param name="mainString">mainString is the string expression from which to extract.</param>
  ''' <param name="matchStr">matchString is the string expression to extract up to. Extract is case-sensitive.</param>
  ''' <returns>Extract returns a sub-string of mainString, starting with its first character (or the character specified by startIndex) and up to (but not including) the first occurrence of matchString. 
  ''' If matchString is not present in mainString, or either string parameter is empty, all of mainString is returned.
  ''' </returns>
  ''' <remarks>Mimics PB's Extract$()</remarks>
  Public Overloads Shared Function Extract(ByVal mainString As String, ByVal matchStr As String) As String
    Return Extract(1, mainString, False, matchStr)
  End Function

  ''' <summary>
  ''' Extract characters from a string up to a character or group of characters.
  ''' </summary>
  ''' <param name="startIndex">startIndex is the starting position to begin extracting.
  ''' If startIndex is not specified, it will start at position 1. If startIndex is zero, or beyond the length of mainString, an empty string is returned. If startIndex is negative, the starting position is counted from right to left: 
  ''' if -1, the search begins at the last character; if -2, the second to last, and so forth. 
  ''' </param>
  ''' <param name="mainString">mainString is the string expression from which to extract.</param>
  ''' <param name="anyMatchChar">If <see langword="true"/>, matchString specifies a list of single characters to be searched for individually, 
  ''' a match on any one of which will cause the extract operation to be performed up to that character.
  ''' </param>
  ''' <param name="matchStr">matchString is the string expression to extract up to. Extract is case-sensitive.</param>
  ''' <returns>Extract returns a sub-string of mainString, starting with its first character (or the character specified by startIndex) and up to (but not including) the first occurrence of matchString. 
  ''' If matchString is not present in mainString, or either string parameter is empty, all of mainString is returned.
  ''' </returns>
  ''' <remarks>Mimics PB's Extract$()</remarks>
  Public Overloads Shared Function Extract(ByVal startIndex As Int32, ByVal mainString As String, ByVal anyMatchChar As Boolean, ByVal matchStr As String) As String

    ' Safe guards
    If startIndex = 0 OrElse startIndex > mainString.Length Then
      Return vbNullString()
    End If

    If mainString.Length < 1 OrElse matchStr.Length < 1 Then
      Return mainString
    End If

    If startIndex > 0 Then
      mainString = Mid(mainString, startIndex)
    End If

    If anyMatchChar = True Then
      If mainString.IndexOfAny(matchStr.ToCharArray) > -1 Then

        Dim start As Int32 = mainString.Length
        For i As Int32 = 0 To matchStr.Length - 1
          If mainString.Contains(matchStr.ElementAt(i)) Then
            If mainString.IndexOf(matchStr.ElementAt(i)) < start Then
              start = mainString.IndexOf(matchStr.ElementAt(i))
            End If
          End If
        Next
        Return Left(mainString, start)
      Else
        Return vbNullString()
      End If

    Else
      Return Extract(startIndex, mainString, matchStr)
    End If

  End Function

  ''' <summary>
  ''' Extract characters from a string up to a character or group of characters.
  ''' </summary>
  ''' <param name="startIndex">startIndex is the starting position to begin extracting.
  ''' If startIndex is not specified, it will start at position 1. If startIndex is zero, or beyond the length of mainString, an empty string is returned. If startIndex is negative, the starting position is counted from right to left: 
  ''' if -1, the search begins at the last character; if -2, the second to last, and so forth. 
  ''' </param>
  ''' <param name="mainString">mainString is the string expression from which to extract.</param>
  ''' <param name="matchStr">matchString is the string expression to extract up to. Extract is case-sensitive.</param>
  ''' <returns>Extract returns a sub-string of mainString, starting with its first character (or the character specified by startIndex) and up to (but not including) the first occurrence of matchString. 
  ''' If matchString is not present in mainString, or either string parameter is empty, all of mainString is returned.
  ''' </returns>
  ''' <remarks>Mimics PB's Extract$()</remarks>
  Public Overloads Shared Function Extract(ByVal startIndex As Int32, ByVal mainString As String, ByVal matchStr As String) As String

    ' Safe guards
    If startIndex = 0 OrElse startIndex > mainString.Length Then
      Return vbNullString()
    End If

    If mainString.Length < 1 OrElse matchStr.Length < 1 Then
      Return mainString
    End If

    If startIndex > 0 Then
      mainString = Mid(mainString, startIndex)
    End If

    If mainString.Contains(matchStr) Then
      Return Mid(mainString, startIndex, InStr(mainString, matchStr) - 1)
    Else
      Return vbNullString()
    End If

  End Function

#End Region

#Region "Chr()"

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

#Region "InStr()"

  ''' <summary>
  ''' Returns an integer specifying the start position of the first occurrence of one string within another.
  ''' </summary>
  ''' <param name="start">
  ''' Numeric expression that sets the starting position for each search. If omitted, search begins at the first character position. The start index is 1-based.
  ''' </param>
  ''' <param name="string1">String expression being searched.</param>
  ''' <param name="string2">String expression sought.</param>
  ''' <returns>
  ''' 0 if <paramref name="string1"/> is zero length or <see langword="null"/>.
  ''' <paramref name="start"/> if <paramref name="string2"/> is zero length or <see langword="null"/>.
  ''' Position where match begins if <paramref name="string2"/> is found in <paramref name="string1"/>.
  ''' 0 if <paramref name="start"/> is > length of <paramref name="string1"/>.
  ''' </returns>
  Public Overloads Shared Function InStr(ByVal start As Int32, ByVal string1 As String, ByVal string2 As String) As Int32

    ' Safe guards
    If start <= 0 Then
      Throw New ArgumentOutOfRangeException("start must be > 0.", "start")
    End If

    If start > string1.Length Then
      Return 0
    End If

    If String.IsNullOrEmpty(string1) Then
      Return 0
    End If

    If String.IsNullOrEmpty(string2) Then
      Return start
    End If

    Return string1.IndexOf(string2, start - 1) + 1

  End Function

  ''' <summary>
  ''' Returns an integer specifying the start position of the first occurrence of one string within another.
  ''' </summary>
  ''' <param name="string1">String expression being searched.</param>
  ''' <param name="string2">String expression sought.</param>
  ''' <returns>
  ''' 0 if <paramref name="string1"/> is zero length or <see langword="null"/>.
  ''' The starting position for the search, which defaults to the first character position if <paramref name="string2"/> is zero length or <see langword="null"/>.
  ''' Position where match begins if <paramref name="string2"/> is found in <paramref name="string1"/>.
  ''' </returns>
  Public Overloads Shared Function InStr(ByVal string1 As String, ByVal string2 As String) As Int32
    Return InStr(1, string1, string2)
  End Function

#End Region

#Region "MCase()"

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

#End Region

#Region "Space()"

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
  Public Overloads Shared Function Space(ByVal count As Int64) As String
    Return New String(" "c, CType(count, Integer))
  End Function

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
  Public Overloads Shared Function Space(ByVal count As UInt64) As String
    Return New String(" "c, CType(count, Integer))
  End Function

#End Region

#Region "String()"

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

#End Region

#Region "Remain()"

  ''' <summary>
  ''' Return the portion of a string following the first occurrence of a character or group of characters.
  ''' </summary>
  ''' <param name="mainString">mainString is searched for the string specified in matchString. If found, all characters after matchString are returned.</param>
  ''' <param name="anyMatchChar">If <see langword="true"/>, matchString specifies a list of single characters to be searched for individually, 
  ''' a match on any one of which will cause the extract operation to be performed up to that character.
  ''' </param>
  ''' <param name="matchStr">matchString is the string expression after which the remainder of mainString is returned. Remain is case-sensitive.</param>
  ''' <returns>Remain returns a sub-string of mainString, following with its first character (or the character specified by startIndex) after (but not including) the first occurrence of matchString. 
  ''' If matchString is not present in mainString, or either string parameter is empty, all of mainString is returned.
  ''' </returns>
  ''' <remarks>Mimics PB's Remain$()</remarks>
  Public Overloads Shared Function Remain(ByVal mainString As String, ByVal anyMatchChar As Boolean, ByVal matchStr As String) As String
    Return Remain(1, mainString, anyMatchChar, matchStr)
  End Function

  ''' <summary>
  ''' Return the portion of a string following the first occurrence of a character or group of characters.
  ''' </summary>
  ''' <param name="mainString">mainString is searched for the string specified in matchString. If found, all characters after matchString are returned.</param>
  ''' <param name="matchStr">matchString is the string expression after which the remainder of mainString is returned. Remain is case-sensitive.</param>
  ''' <returns>Remain returns a sub-string of mainString, following with its first character (or the character specified by startIndex) after (but not including) the first occurrence of matchString. 
  ''' If matchString is not present in mainString, or either string parameter is empty, all of mainString is returned.
  ''' </returns>
  ''' <remarks>Mimics PB's Remain$()</remarks>
  Public Overloads Shared Function Remain(ByVal mainString As String, ByVal matchStr As String) As String
    Return Remain(1, mainString, False, matchStr)
  End Function

  ''' <summary>
  ''' Return the portion of a string following the first occurrence of a character or group of characters.
  ''' </summary>
  ''' <param name="startIndex">startIndex is the starting position to begin extracting.
  ''' If startIndex is not specified, it will start at position 1. If startIndex is zero, or beyond the length of mainString, an empty string is returned. If startIndex is negative, the starting position is counted from right to left: 
  ''' if -1, the search begins at the last character; if -2, the second to last, and so forth. 
  ''' </param>
  ''' <param name="mainString">mainString is searched for the string specified in matchString. If found, all characters after matchString are returned.</param>
  ''' <param name="anyMatchChar">If <see langword="true"/>, matchString specifies a list of single characters to be searched for individually, 
  ''' a match on any one of which will cause the extract operation to be performed up to that character.
  ''' </param>
  ''' <param name="matchStr">matchString is the string expression after which the remainder of mainString is returned. Remain is case-sensitive.</param>
  ''' <returns>Remain returns a sub-string of mainString, following with its first character (or the character specified by startIndex) after (but not including) the first occurrence of matchString. 
  ''' If matchString is not present in mainString, or either string parameter is empty, all of mainString is returned.
  ''' </returns>
  ''' <remarks>Mimics PB's Remain$()</remarks>
  Public Overloads Shared Function Remain(ByVal startIndex As Int32, ByVal mainString As String, ByVal anyMatchChar As Boolean, ByVal matchStr As String) As String

    ' Safe guards
    If startIndex = 0 OrElse startIndex > mainString.Length Then
      Return vbNullString()
    End If

    If mainString.Length < 1 OrElse matchStr.Length < 1 Then
      Return mainString
    End If

    If startIndex > 0 Then
      mainString = Mid(mainString, startIndex)
    End If

    If anyMatchChar = True Then
      If mainString.IndexOfAny(matchStr.ToCharArray) > -1 Then

        Dim start As Int32 = mainString.Length
        For i As Int32 = 0 To matchStr.Length - 1
          If mainString.Contains(matchStr.ElementAt(i)) Then
            If mainString.IndexOf(matchStr.ElementAt(i)) < start Then
              start = mainString.IndexOf(matchStr.ElementAt(i))
            End If
          End If
        Next
        Return Mid(mainString, start + 2)
      Else
        Return vbNullString()
      End If

    Else
      Return Remain(startIndex, mainString, matchStr)
    End If

  End Function

  ''' <summary>
  ''' Return the portion of a string following the first occurrence of a character or group of characters.
  ''' </summary>
  ''' <param name="startIndex">startIndex is the starting position to begin extracting.
  ''' If startIndex is not specified, it will start at position 1. If startIndex is zero, or beyond the length of mainString, an empty string is returned. If startIndex is negative, the starting position is counted from right to left: 
  ''' if -1, the search begins at the last character; if -2, the second to last, and so forth. 
  ''' </param>
  ''' <param name="mainString">mainString is searched for the string specified in matchString. If found, all characters after matchString are returned.</param>
  ''' <param name="matchStr">matchString is the string expression after which the remainder of mainString is returned. Remain is case-sensitive.</param>
  ''' <returns>Remain returns a sub-string of mainString, following with its first character (or the character specified by startIndex) after (but not including) the first occurrence of matchString. 
  ''' If matchString is not present in mainString, or either string parameter is empty, all of mainString is returned.
  ''' </returns>
  ''' <remarks>Mimics PB's Remain$()</remarks>
  Public Overloads Shared Function Remain(ByVal startIndex As Int32, ByVal mainString As String, ByVal matchStr As String) As String

    ' Safe guards
    If startIndex = 0 OrElse startIndex > mainString.Length Then
      Return vbNullString()
    End If

    If mainString.Length < 1 OrElse matchStr.Length < 1 Then
      Return mainString
    End If

    If startIndex > 0 Then
      mainString = Mid(mainString, startIndex)
    End If

    If mainString.Contains(matchStr) Then
      Return Mid(mainString, InStr(mainString, matchStr) + matchStr.Length)
    Else
      Return vbNullString()
    End If

  End Function

#End Region

#Region "StrIncrement()"

  ''' <summary>
  ''' "Increments" a string, e.g. takes the character and numerical portion of a string and increments it.
  ''' </summary>
  ''' <param name="mainString">Increment this string.</param>
  ''' <returns>Incremented mainString</returns>
  ''' <remarks>
  ''' Numbers and characters become incremented naturally, i.e. "5" -> "6", "a" -> "c".
  ''' The string as a whole will be "incremented" following this rules, e.g.
  ''' "aa" -> "ab"
  ''' "a0" -> "a1"
  ''' The interesting part happens if "9"  -> "0" and "z" -> "a", e.g. "a9z" -> "b0a"
  ''' The logic behind this is z "plus 1" becomes a, but just as 9 plus 1 becomes 10,
  ''' we need to carry over the "1" of the "10" to the next higher base, so "a" becomes "b".
  ''' Any non-alphanumeric character in <paramref name="mainString"/> remains as is, e.g.
  ''' 00-AA-Z9 -> 00-AB-A0
  ''' </remarks>
  Public Shared Function StrIncrement(ByVal mainString As String) As String

    'StrIncr procedure for PowerBASIC
    ' by Dave Navarro, Jr.
    ' Donated to the Public Domain
    ' Last Revision: July 15, 1994
    ' Update by Knuth Konrad


    Dim lValue, x As Int32
    Dim sChar, sTemp As String

    ' "9"'s only?
    sTemp = mainString.Replace("9", String.Empty)
    If sTemp.Length = 0 Then
      lValue = Convert.ToInt32(mainString)
      lValue = lValue + 1
      Return lValue.ToString
    End If

    ' If TALLY(sString, "9") = LEN(sString) Then
    'String besteht nur aus Zahlen und nur aus
    '9ern
    ' lValue = VAL(sString)
    ' INCR lValue
    ' StrIncr = TRIM$(STR$(lValue))
    ' Exit Function
    ' End If

    For x = mainString.Length To 1 Step -1

      'For x = LEN(sString) To 1 Step -1

      sChar = Mid(mainString, x, 1)

      If (sChar >= "0") And (sChar <= "8") Then
        sChar = Chr(Asc(sChar) + 1)
        Mid(mainString, x, 1) = sChar
        Exit For
      ElseIf sChar = "9" Then
        sChar = "0"
        Mid(mainString, x, 1) = sChar
      ElseIf (sChar >= "A") And (sChar <= "Y") Then
        sChar = Chr(Asc(sChar) + 1)
        Mid(mainString, x, 1) = sChar
        Exit For
      ElseIf sChar = "Z" Then
        sChar = "A"
        Mid(mainString, x, 1) = sChar
      ElseIf (sChar >= "a") And (sChar <= "y") Then
        sChar = Chr(Asc(sChar) + 1)
        Mid(mainString, x, 1) = sChar
        Exit For
      ElseIf sChar = "z" Then
        sChar = "a"
        Mid(mainString, x, 1) = sChar
      End If

    Next x

    Return mainString

  End Function

#End Region

#Region "TrimAny()"

  ''' <summary>
  ''' Remove any occurrence of <paramref name="removeChars"/> in <paramref name="source"/>
  ''' </summary>
  ''' <param name="source">Source string</param>
  ''' <param name="removeChars">List of strings to remove from <paramref name="source"/></param>
  ''' <returns>Trimmed string.</returns>
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
  ''' <returns>Trimmed string.</returns>
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
  ''' <returns>Trimmed string.</returns>
  Public Overloads Shared Function TrimAny(ByVal source As String, ByVal removeChars() As String) As String

    Dim result As String = source

    For Each s As String In removeChars
      result = result.TrimStart(s.ToCharArray)
      result = result.TrimEnd(s.ToCharArray)
    Next

    Return result

  End Function

#End Region

  ''' <summary>
  ''' Encloses <paramref name="text"/> with double quotation marks (").
  ''' </summary>
  ''' <param name="text">Wrap this string in quotation marks.</param>
  ''' <returns><paramref name="text"/> enclosed in double quotation marks (")</returns>
  Public Shared Function EnQuote(ByVal text As String) As String
    Return System.Convert.ToChar(34).ToString & text & System.Convert.ToChar(34).ToString
  End Function

#Region "Left()"
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
#End Region

#Region "Mid()"
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
#End Region

#Region "Right()"
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
#End Region

#Region "Wrap()"

  ''' <summary>
  ''' Wraps a string in double quotes (")/>
  ''' </summary>
  ''' <param name="text">Original string</param>
  ''' <returns>Enclosed <paramref name="text"/></returns>
  Public Shared Function Wrap(ByVal text As String) As String
    Return Wrap(text, vbQuote)
  End Function

  ''' <summary>
  ''' Wraps a string into <paramref name="leftChar"/> And <paramref name="rightChar"/>
  ''' </summary>
  ''' <param name="text">Original string</param>
  ''' <param name="leftChar">Left 'bracket'</param>
  ''' <param name="rightChar">Right 'bracket'. If <paramref name="rightChar"/> = "", <paramref name="leftChar"/> is used instead.</param>
  ''' <returns>Enclosed <paramref name="text"/></returns>
  Public Shared Function Wrap(ByVal text As String, ByVal leftChar As String, Optional ByVal rightChar As String = "") As String

    If rightChar.Length < 1 Then
      Return leftChar & text & leftChar
    Else
      Return leftChar & text & rightChar
    End If

  End Function

  ''' <summary>
  ''' Wraps a string into <paramref name="wrapChar"/>
  ''' </summary>
  ''' <param name="text">Original string</param>
  ''' <param name="wrapChar">'Bracket' character</param>
  ''' <returns>Enclosed <paramref name="text"/></returns>
  Public Shared Function Wrap(ByVal text As String, ByVal wrapChar As Char) As String
    Return wrapChar & text & wrapChar
  End Function

#End Region

#Region "UnWrap()"

  ''' <summary>
  ''' Unwraps a string from double quotes.
  ''' </summary>
  ''' <param name="text">Original string</param>
  ''' <returns><paramref name="text"/> w/o outer double quotes, e.g. "(My text)" becomes (My text)</returns>
  Public Shared Function UnWrap(ByVal text As String) As String

    ' Safe guard
    If (text.Length < 1) Then
      Return text
    Else
      Return UnWrap(text, vbQuote)
    End If

  End Function

  ''' <summary>
  ''' Unwraps a string from <paramref name="leftChar"/> And <paramref name="rightChar"/>
  ''' </summary>
  ''' <param name="text">Original string</param>
  ''' <param name="leftChar">Left 'bracket'</param>
  ''' <param name="rightChar">Right 'bracket'. If <paramref name="rightChar"/> = "", <paramref name="leftChar"/> is used instead.</param>
  ''' <returns>Enclosed <paramref name="text"/></returns>
  Public Shared Function UnWrap(ByVal text As String, ByVal leftChar As String, Optional ByVal rightChar As String = "") As String

    ' Safe guard
    If (text.Length < 1) Then
      Return text
    End If

    If (text.Length > leftChar.Length And text.StartsWith(leftChar)) Then
      text = text.Substring(leftChar.Length)
    End If

    If (rightChar.Length < 1) Then
      rightChar = leftChar
    End If

    If (text.Length > leftChar.Length And text.EndsWith(leftChar)) Then
      text = text.Substring(0, text.Length - leftChar.Length)
    End If

    Return text

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
  ''' Constant for a double quotation mark (").
  ''' </summary>
  ''' <param name="n">Number of DQs to return</param>
  ''' <returns><paramref name="n"/> "</returns>
  Public Shared Function vbDoubleQuote(Optional ByVal n As Int32 = 1) As String
    Return vbQuote(n)
  End Function

  ''' <summary>
  ''' Constant for a single quotation mark/apostrophe (').
  ''' </summary>
  ''' <param name="n">Number of SQs to return</param>
  ''' <returns><paramref name="n"/> "</returns>
  Public Shared Function vbSingleQuote(Optional ByVal n As Int32 = 1) As String
    Dim sResult As String = String.Empty
    For i As Int32 = 1 To n
      sResult &= System.Convert.ToChar(39).ToString
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
  ''' <returns>'White space'</returns>
  Public Shared Function vbWhiteSpace() As String
    Return vbTab() & vbNewLine() & " "
  End Function

#End Region

#Region "ParseCount"

  ''' <summary>
  ''' Return the count of (comma) delimited strings in a string expression.
  ''' </summary>
  ''' <param name="stringExpression">
  ''' This is the string to examine and parse. If <paramref name="stringExpression"/> is empty (a null string) or contains no delimiter character(s), the string is considered to contain exactly one sub-field. 
  ''' In this case, ParseCount returns the value 1.
  '''</param>
  ''' <returns>Number of strings in <paramref name="stringExpression"/></returns>
  ''' <remarks>Mimics PB's ParseCount</remarks>
  Public Shared Function ParseCount(
    ByVal stringExpression As String
    ) As Int32

    Dim count As Int32

    ' Safe guard
    If System.String.IsNullOrEmpty(stringExpression) OrElse Not stringExpression.Contains(",") Then
      count = 1

    Else

      Try
        Dim tmp As String() = stringExpression.Split(",")
        count = tmp.Count
      Catch ex As Exception
        count = 0
      End Try

    End If

    Return count

  End Function

  ''' <summary>
  ''' Return the count of (comma) delimited strings in a string expression.
  ''' </summary>
  ''' <param name="stringExpression">
  ''' This is the string to examine and parse. If <paramref name="stringExpression"/> is empty (a null string) or contains no delimiter character(s), the string is considered to contain exactly one sub-field. 
  ''' In this case, ParseCount returns the value 1.
  ''' </param>
  ''' <param name="delimiters">
  ''' Characters considered to be delimiters
  ''' </param>
  ''' <returns>Number of strings in <paramref name="stringExpression"/></returns>
  ''' <remarks>Mimics PB's ParseCount</remarks>
  Public Shared Function ParseCount(
    ByVal stringExpression As String,
    ByVal delimiters As String
    ) As Int32

    Dim count As Int32

    ' Safe guard
    If System.String.IsNullOrEmpty(stringExpression) OrElse System.String.IsNullOrEmpty(delimiters) Then
      count = 1
    Else

      For i As Int32 = 1 To delimiters.Length

        If stringExpression.Contains(Mid(delimiters, i, 1)) Then
          Try
            Dim tmp As String() = stringExpression.Split(Mid(delimiters, i, 1))
            count += tmp.Count
          Catch ex As Exception
            count += 0
          End Try
        End If

      Next i

      ' stringExpression isn't empty, but no delimiters found = return one
      If count = 0 Then
        count += 1
      End If
    End If

    Return count

  End Function

#End Region

End Class
