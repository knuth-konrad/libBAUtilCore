Imports System.IO

Imports System.Runtime.Serialization.Formatters.Binary

Imports System.Xml
Imports System.Xml.Serialization

''' <summary>
''' General object helpers
''' </summary>
Public Class ObjectHelper

#Region "Serialization"
   ''' <summary>
   ''' Returns the enumeration member's name for the specific enumeration value.
   ''' </summary>
   ''' <param name="enumType">.NET type of enum as retrieved by <see cref="Type.GetType"/>.</param>
   ''' <param name="enumMemberValue">Return the member name for this value</param>
   ''' <returns>
   ''' The enumeration's member name matching <paramref name="enumMemberValue"/>
   ''' </returns>
   ''' <remarks>
   ''' <paramref name="enumType"/> MUST be an Enumeration's *type*, e.g. GetType(MyEnum)
   ''' </remarks>
   Public Overloads Shared Function GetEnumNameFromValue(ByVal enumType As Type, ByVal enumMemberValue As Int32) As String
      '------------------------------------------------------------------------------
      'Prereq.  : 
      '
      '   Author: Knuth Konrad
      '     Date: 30.08.2018
      '   Source: Modified from https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/enum-statement
      '  Changed: 2019-01-12
      '           - Enum member names naturally starting with a number should be prefixed 
      '           with a "_", e.g. _12_Passenger_Van. The underscore is then removed
      '------------------------------------------------------------------------------
      Dim names = [Enum].GetNames(enumType)
      Dim values = [Enum].GetValues(enumType)
      Dim sTemp As String = String.Empty

      For i As Int32 = 0 To names.Length - 1
         If CType(values.GetValue(i), Int32) = enumMemberValue Then
            sTemp = names(i)
            If sTemp.StartsWith("_") = True Then
               sTemp = sTemp.Substring(1)
            End If
            Return sTemp
         End If
      Next

      Return String.Empty

   End Function

   ''' <summary>
   ''' Returns the enumeration member's name for the specific enumeration value.
   ''' </summary>
   ''' <param name="enumType">.NET type of enum as retrieved by <see cref="Type.GetType"/>.</param>
   ''' <param name="enumMemberValue">Return the member name for this value.</param>
   ''' <param name="alternativeNames">
   ''' Array with alternative names to return.
   ''' Empty array elements of alternativeNames() will cause the Enum's member name to be returned. E.g.
   ''' Enum MyEnum
   '''    One
   '''    Two
   '''    Three
   ''' End Enum
   ''' alternativeNames() = "", "", "more than two"
   ''' will return "One" for enumMemberValue = MyEnum.One, "Two" for MyEnum.Two, but "more than two" for MyEnum.Three
   ''' </param>
   ''' <returns>
   ''' The enumeration's member name matching <paramref name="enumMemberValue"/>.
   ''' </returns>
   ''' <remarks>
   ''' <paramref name="enumType"/> MUST be an Enumeration's *type*, e.g. GetType(MyEnum)
   ''' </remarks>
   Public Overloads Shared Function GetEnumNameFromValue(ByVal enumType As Type, ByVal enumMemberValue As Int32,
                                                         ByVal ParamArray alternativeNames() As String) As String
      '------------------------------------------------------------------------------
      'Prereq.  : enumType MUST be an Enumeration's *type*, e.g. GetType(MyEnum)
      '
      '   Author: Knuth Konrad
      '     Date: 24.04.2019
      '   Source: Modified from https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/enum-statement
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim names = [Enum].GetNames(enumType)
      Dim values = [Enum].GetValues(enumType)
      Dim sTemp As String = String.Empty

      ' Safe guard
      If alternativeNames.Length <> names.Length Then
         Throw New ArgumentOutOfRangeException("alternativeNames(). The number of array elements must match the number of Enumeration members.")
      End If

      For i As Int32 = 0 To names.Length - 1
         If CType(values.GetValue(i), Int32) = enumMemberValue Then
            If Not String.IsNullOrEmpty(alternativeNames(i)) Then
               sTemp = alternativeNames(i)
            Else
               sTemp = names(i)
            End If
            If sTemp.StartsWith("_") = True Then
               sTemp = sTemp.Substring(1)
            End If
            Return sTemp
         End If
      Next

      Return String.Empty

   End Function

   ''' <summary>
   ''' Determine if an object is serializable.
   ''' </summary>
   ''' <param name="obj">Check this object</param>
   ''' <returns><see langword="true"/> if <paramref name="obj"/> can be serualized, <see langword="false"/> otherwise.</returns>
   Public Shared Function IsSerializable(ByVal obj As Object) As Boolean
      '------------------------------------------------------------------------------
      'Prereq.  : -
      '
      '   Author: Knuth Konrad
      '     Date: 28.05.2019
      '   Source: https://docs.microsoft.com/en-us/dotnet/standard/serialization/how-to-determine-if-netstandard-object-is-serializable
      '  Changed: -
      '------------------------------------------------------------------------------

      ' Safe guard
      If obj Is Nothing Then
         Return False
      Else
         Dim t As Type = obj.GetType()
         Return t.IsSerializable
      End If

   End Function

   ''' <summary>
   ''' Serialize an object to (a) XML (string).
   ''' </summary>
   ''' <param name="obj">Serialize this object</param>
   ''' <param name="omitXmlDeclaration">True = Serialize without XML declaration (&lt;?xml version="1.0"?&gt;)</param>
   ''' <param name="omitXmlNamespace">
   ''' True = Serialize without XML namespace declaration
   ''' (xmlnsxsi = "http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd = "http://www.w3.org/2001/XMLSchema")
   ''' </param>
   ''' <returns>Serialized <paramref name="obj"/> as an (XML) string.</returns>
   Public Shared Function Serialize(ByVal obj As Object, Optional ByVal omitXmlDeclaration As Boolean = False,
                                    Optional ByVal omitXmlNamespace As Boolean = False) As String
      '------------------------------------------------------------------------------
      'Prereq.  : -
      '
      '   Author: Knuth Konrad
      '     Date: 28.05.2019
      '   Source: https://www.it-visions.de/lserver/CodeSampleDetails.aspx?c=2831
      '           https://stackoverflow.com/questions/258960/how-to-serialize-an-object-to-xml-without-getting-xmlns
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim serializer As New XmlSerializer(obj.GetType)
      Dim settings As New XmlWriterSettings()
      Dim ns As New XmlSerializerNamespaces()
      Dim s As String = String.Empty

      If omitXmlNamespace = True Then
         ' Create an empty namespace
         ns.Add("", "")
      End If

      'If omitXmlDeclaration = True Then
      '   settings.OmitXmlDeclaration = True
      'End If
      settings.OmitXmlDeclaration = omitXmlDeclaration

      Dim ms As New MemoryStream()
      Dim sw As XmlWriter = XmlWriter.Create(ms, settings)
      Dim sr As New StreamReader(ms)

      serializer.Serialize(sw, obj, ns)
      ms.Position = 0

      Return sr.ReadToEnd()

   End Function

   ''' <summary>
   ''' Deserialize a XML (string) to an object.
   ''' </summary>
   ''' <param name="xmlString">XML as String compatible with .NET's Deserialize method.</param>
   ''' <param name="objType">Deserialize to this object type.</param>
   ''' <returns>Object of type <paramref name="objType"/>.</returns>
   Public Overloads Shared Function Deserialize(ByVal xmlString As String, ByVal objType As Type) As Object
      '------------------------------------------------------------------------------
      'Prereq.  : -
      'Note     : -
      '
      '   Author: Knuth Konrad
      '     Date: 29.05.2019
      '   Source: https://stackoverflow.com/questions/27235951/deserialize-xml-string-to-object-vb-net
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim xser As New XmlSerializer(objType)
      Dim sr As TextReader = New StringReader(xmlString)

      Return xser.Deserialize(sr)

   End Function

   ''' <summary>
   ''' Deserialize a XML (string) to a specific class.
   ''' </summary>
   ''' <typeparam name="T">Return an object of this class.</typeparam>
   ''' <param name="xmlString">XML as String compatible with .NET's Deserialize method.</param>
   ''' <returns>Object of type <typeparamref name="T"/>.</returns>
   Public Shared Function DeserializeAsClass(Of T As Class)(ByVal xmlString As String) As T
      '------------------------------------------------------------------------------
      'Prereq.  : -
      'Note     : -
      '
      '   Author: Knuth Konrad
      '     Date: 29.05.2019
      '   Source: https://www.codeguru.com/csharp/.net/net_data/serializing-and-deserializing-xml-in-.net.html
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim xser As XmlSerializer = New XmlSerializer(GetType(T))
      Dim sr As StringReader = New StringReader(xmlString)

      Return CType(xser.Deserialize(sr), T)

   End Function
#End Region

   ''' <summary>
   ''' Create a deep copy of an object
   ''' </summary>
   ''' <param name="obj">Object to be cloned</param>
   ''' <returns>
   ''' Deep copy aka "a copy of all data and objects and their data" of <paramref name="obj"/>.
   ''' </returns>
   ''' <remarks>Source: https://www.rectanglered.com/deep-copying-object-vb-net </remarks>
   Public Shared Function Clone(ByVal obj As Object) As Object

      Dim m As New MemoryStream()
      Dim f As New BinaryFormatter()

      f.Serialize(m, obj)
      m.Seek(0, SeekOrigin.Begin)

      Return f.Deserialize(m)

   End Function

   ''' <summary>
   ''' Simple wrapper to <see cref="Clone(Object)"/>
   ''' </summary>
   ''' <param name="o"></param>
   ''' <returns>
   ''' Deep copy aka "a copy of all data and objects and their data" of <paramref name="o"/>.
   ''' </returns>
   Public Shared Function DeepClone(ByVal o As Object) As Object
      Return Clone(o)
   End Function

End Class
