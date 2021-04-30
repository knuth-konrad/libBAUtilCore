Imports System.ComponentModel

''' <summary>
''' General purpose helper methods
''' </summary>
Public Class baUtil

#Region "Declares"

   ''' <summary>
   '''    Mimic Microsoft.VisualBasic.TriState
   ''' </summary>
   Public Enum TriState
      [False] = 0
      [True] = -1
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
   Public Function ToLower() As String
      Return Me.Value.ToLower
   End Function

   ''' <summary>
   ''' Return the uppercase representation of the string's content
   ''' </summary>
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
