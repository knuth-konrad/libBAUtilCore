Imports System.ComponentModel

''' <summary>
''' General multipurpose utility library.
''' </summary>
''' <remarks>
''' Generates an XML comments for namespaces for SandCastle.
''' See https://stackoverflow.com/questions/793210/xml-documentation-for-a-namespace
''' </remarks>
<System.Runtime.CompilerServices.CompilerGeneratedAttribute()>
Class NameSpaceDoc
End Class

''' <summary>
''' General purpose helper methods
''' </summary>
Public Class baUtil

#Region "Declares"

  ''' <summary>
  ''' Mimic Microsoft.VisualBasic.TriState
  ''' </summary>
  Public Enum TriState
    ''' <summary>
    ''' VB False
    ''' </summary>
    [False] = 0
    ''' <summary>
    ''' VB True
    ''' </summary>
    [True] = -1
    ''' <summary>
    ''' VB UseDefault
    ''' </summary>
    [UseDefault] = -2
  End Enum

#End Region

#Region "Formatting (Strings / Numbers)"

  'Public Shared Function FormatNumber(ByVal Expression As Object,
  '                                    Optional ByVal NumDigitsAfterDecimal As Int32 = -1,
  '                                    Optional ByVal IncludeLeadingDigit As TriState = TriState.UseDefault,
  '                                    Optional ByVal UseParensForNegativeNumbers As TriState = TriState.UseDefault,
  '                                    Optional ByVal GroupDigits As TriState = TriState.UseDefault) As String

  '   Dim sFormatMask As String = String.Empty
  '   Dim curCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CurrentCulture

  '   Console.WriteLine(Expression.ToString)

  '   If NumDigitsAfterDecimal = -1 Then
  '      ' -1 = use system defaults

  '   End If

  'End Function

#End Region

#Region "Sleep"
  ''' <summary>
  ''' Mimics VB6 Sleep
  ''' </summary>
  ''' <param name="milliSeconds"></param>
  Public Shared Sub Sleep(ByVal milliSeconds As Int32)
    System.Threading.Thread.Sleep(milliSeconds)
  End Sub
#End Region

End Class

''' <summary>
''' General math helpers
''' </summary>
Public Class MathUtil

  ''' <summary>
  ''' Returns the % of Total given by Part, e.g. Total = 200, Part = 50 = 25(%)
  ''' </summary>
  ''' <param name="part">Part of <paramref name="total"/> to be expressed as a percent value.</param>
  ''' <param name="total">Value considered to be 100%.</param>
  ''' <returns><paramref name="part"/> percent of <paramref name="total"/></returns>
  Public Shared Function Percent(ByVal part As Double, ByVal total As Double) As Double

    If total = 0 Then
      Throw New System.ArgumentOutOfRangeException("total", "Value can't be zero.")
    End If

    Return (part / total) * 100

  End Function

  ''' <summary>
  ''' Mimics VB6's Mod operator behavior which converts both operants to integer first, then 
  ''' does the division, wheres VB.NET divides the (floating point) numbers and then does the 
  ''' integer conversion
  ''' </summary>
  ''' <param name="operand1">Divide this by <paramref name="operand2"/>.</param>
  ''' <param name="operand2">Divide <paramref name="operand1"/> by this.</param>
  ''' <returns><paramref name="operand2"/>Modulo division result</returns>
  ''' <remarks>Source:
  ''' https://stackoverflow.com/questions/73362557/what-does-upgrade-warning-mod-has-a-new-behavior-refer-to
  ''' </remarks>
  Public Shared Function ModVB6(ByVal operand1 As Double, ByVal operand2 As Double) As Int32

    If (operand2 = 0) OrElse (CType(operand2, Int32) = 0) Then
      Throw New System.ArgumentOutOfRangeException("operand2", "Value can't be zero.")
    End If

    Return (CType(operand1, Int32) Mod CType(operand2, Int32))

  End Function

End Class

''' <summary>
''' Mimics VB6's fixed string
''' </summary>
<DefaultProperty("Value")>
<Serializable()> Public Class FixedLengthString

  Private Const PADDING_CHAR As Char = " "c

  Private miLength As UInt16
  Private miMinLength As UInt16
  Private mcPaddingChar As Char = PADDING_CHAR
  Private msValue As String

  ''' <summary>
  ''' Fixed length of the string
  ''' </summary>
  Public Property Length As UInt16
    Get
      Return miLength
    End Get
    Set(value As UInt16)
      miLength = value
    End Set
  End Property

  ''' <summary>
  ''' Minimum length
  ''' </summary>
  Public Property MinLength As UInt16
    Get
      Return miMinLength
    End Get
    Set(value As UInt16)
      miMinLength = value
    End Set
  End Property

  ''' <summary>
  ''' String content
  ''' </summary>
  Public Property Value As String
    Get
      Return msValue
    End Get
    Set(value As String)
      With Me
        If (.MinLength > 0) And (value.Length < .MinLength) Then
          Throw New ArgumentOutOfRangeException("Value", "Length < MinLength")
        End If

        If value.Length < .MinLength Then
          msValue = value & New String(CType(" ", Char), .MinLength - value.Length)
        Else
          msValue = value.Substring(0, .MinLength)
        End If
      End With
    End Set
  End Property

  ''' <summary>
  ''' Padding character used to fill unoccupied string space
  ''' </summary>
  Public Property PaddingChar As Char
    Get
      Return mcPaddingChar
    End Get
    Set(value As Char)
      mcPaddingChar = value
    End Set
  End Property

  ''' <summary>
  ''' Return the lowercase representation of the string's content
  ''' </summary>
  ''' <returns>A string in lowercase.</returns>
  Public Function ToLower() As String
    Return Me.Value.ToLower
  End Function

  ''' <summary>
  ''' Return the uppercase representation of the string's content
  ''' </summary>
  ''' <returns>A string in uppercase.</returns>
  Public Function ToUpper() As String
    Return Me.Value.ToUpper
  End Function

  ''' <summary>
  ''' Initializes a new instance of the FixedLengthString class
  ''' </summary>
  ''' <param name="stringLength">Length of string</param>
  Public Sub New(ByVal stringLength As UInt16)

    MyBase.New
    With Me
      .Length = stringLength
      .Value = New String(.PaddingChar, stringLength)
    End With
  End Sub

  ''' <summary>
  ''' Initializes a new instance of the FixedLengthString class
  ''' </summary>
  ''' <param name="stringLength">Length of string</param>
  ''' <param name="stringValue">String content</param>
  Public Sub New(ByVal stringLength As UInt16, ByVal stringValue As String)

    MyBase.New
    With Me
      .Length = stringLength
      .Value = stringValue
      If stringValue.Length < .Length Then
        .Value = stringValue.PadRight(.Length, .PaddingChar)
      End If
    End With

  End Sub

  ''' <summary>
  ''' Initializes a new instance of the FixedLengthString class
  ''' </summary>
  ''' <param name="stringLength">Length of string</param>
  ''' <param name="stringMinLength">Minimum length of string</param>
  Public Sub New(ByVal stringLength As UInt16, ByVal stringMinLength As UInt16)

    MyBase.New
    If (stringMinLength > 0) And (stringLength < stringMinLength) Then
      Throw New ArgumentOutOfRangeException("Length", "Length can't be < MinLength")
    End If

    With Me
      .Length = stringLength
      .MinLength = stringMinLength
      .Value = New String(.PaddingChar, stringLength)
    End With

  End Sub

  ''' <summary>
  ''' Initializes a new instance of the FixedLengthString class
  ''' </summary>
  ''' <param name="stringLength">Length of string</param>
  ''' <param name="stringValue">String content</param>
  ''' <param name="stringMinLength">Minimum length of string</param>
  Public Sub New(ByVal stringLength As UInt16, ByVal stringValue As String, ByVal stringMinLength As UInt16)

    MyBase.New
    If (stringMinLength > 0) And (stringLength < stringMinLength) Then
      Throw New ArgumentOutOfRangeException("Length", "Length can't be < MinLength")
    End If

    With Me
      .Length = stringLength
      .MinLength = stringMinLength
      .Value = stringValue
      If stringValue.Length < .Length Then
        .Value = stringValue.PadRight(.Length, .PaddingChar)
      End If
    End With

  End Sub

End Class
