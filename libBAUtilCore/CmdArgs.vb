'------------------------------------------------------------------------------
'   Author: Knuth Konrad 2021-04-05
'  Changed: -
'------------------------------------------------------------------------------
Imports libBAUtilCore.StringHelper

Namespace Utils.Args

   ''' <summary>
   ''' Command line parameter handling
   ''' See https://docs.microsoft.com/en-us/archive/msdn-magazine/2019/march/net-parse-the-command-line-with-system-commandline
   ''' See https://github.com/commandlineparser/commandline
   ''' </summary>
   Public Class CmdArgs

      Implements System.IDisposable

#Region "Declarations"

      ''' <summary>
      ''' OS-dependent default parameter delimiter.
      ''' </summary>
      Public Enum eArgumentDelimiterStyle
         ''' <summary>
         ''' Windows style parameter delimiter '/'
         ''' </summary>
         Windows
         ''' <summary>
         ''' *nix style parameter delimiter '--'
         ''' </summary>
         POSIX
      End Enum


      ' Default arguments and key/value delimiter
      ''' <summary>
      ''' Standard Windows parameter delimiter.
      ''' </summary>
      Private Const DELIMITER_ARGS_WIN As String = "/"
      ''' <summary>
      ''' Standard POSIX ("Linux") parameter delimiter.
      ''' </summary>
      Private Const DELIMITER_ARGS_POSIX As String = "--"
      ''' <summary>
      ''' Default key=value delimiter.
      ''' </summary>
      Private Const DELIMITER_VALUE As String = "="

      Dim mbolCaseSensitive As Boolean                   ' Treat parameter names as case-sensitive?
      Dim msDelimiterArgs As String = String.Empty       ' Arguments delimiter, typically "/"
      Dim msDelimiterValue As String = String.Empty      ' Key/value delimiter, typically "="
      Dim msOriginalParameters As String = String.Empty  ' Parameters as passed to the application
      Dim listValidParameters As List(Of String)         ' List of all valid parameter

      Private mcolKeyValues As List(Of KeyValue)

#End Region

#Region "Properties - Public"

      ''' <summary>
      ''' Return the number of parameters (key/value)
      ''' </summary>
      Public ReadOnly Property ParametersCount As Int32
         Get
            If Not Me.KeyValues Is Nothing Then
               Return Me.KeyValues.Count
            Else
               Return 0
            End If
         End Get
      End Property

      ''' <summary>
      ''' Treat parameter names as case-sensitive?
      ''' </summary>
      Public Property CaseSensitive As Boolean
         Get
            Return mbolCaseSensitive
         End Get
         Set(value As Boolean)
            mbolCaseSensitive = value
         End Set
      End Property

      ''' <summary>
      ''' Parameter delimiter, e.g. '--' in '--param=value'
      ''' </summary>
      Public Property DelimiterArgs As String
         Get
            Return msDelimiterArgs
         End Get
         Set(value As String)
            msDelimiterArgs = value
         End Set
      End Property

      ''' <summary>
      ''' Name/value delimiter, e.g. '=' in  '--param=value'
      ''' </summary>
      Public Property DelimiterValue As String
         Get
            Return msDelimiterValue
         End Get
         Set(value As String)
            msDelimiterValue = value
         End Set
      End Property

      Public Property ValidParameters As List(Of String)
         Get
            Return listValidParameters
         End Get
         Set(value As List(Of String))
            listValidParameters = value
         End Set
      End Property

      ''' <summary>
      ''' Unmodified parameter string as passed to the application
      ''' </summary>
      Public Property OriginalPrameters As String
         Get
            Return msOriginalParameters
         End Get
         Set(value As String)
            msOriginalParameters = value
         End Set
      End Property

      ''' <summary>
      ''' The parsed command line parameters
      ''' </summary>
      Public Property KeyValues As List(Of KeyValue)
         Get
            Return mcolKeyValues
         End Get
         Set(value As List(Of KeyValue))
            mcolKeyValues = value
         End Set
      End Property

#End Region

#Region "Methods - Private"

      ''' <summary>
      ''' Retrieve the parameter delimiter according to the OS' typical flavor
      ''' </summary>
      ''' <returns>OS typical parameter delimiter</returns>
      Private Function GetDefaultDelimiterForOS() As String

         If System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows) Then
            Return DELIMITER_ARGS_WIN
         Else
            Return DELIMITER_ARGS_POSIX
         End If

      End Function

      ''' <summary>
      ''' Is a parameter present more than once?
      ''' </summary>
      Private Function HasDuplicate(ByVal currentKey As KeyValue, ByVal currentIndex As Int32) As Boolean

         ' Safe guard
         If Me.KeyValues.Count < 1 Then
            Return False
         End If

         ' Compare the passed parameter to the list and see if there's a duplicate entry
         With Me
            For i As Int32 = 0 To .KeyValues.Count - 1
               Dim o As KeyValue = .KeyValues.Item(i)
               If .CaseSensitive = False Then
                  If (o.Key.ToLower = currentKey.Key.ToLower) AndAlso (i <> currentIndex) Then
                     Return True
                  End If
               Else
                  If (o.Key = currentKey.Key) AndAlso (i <> currentIndex) Then
                     Return True
                  End If
               End If
            Next
         End With

         ' Reaching here, we've found no duplicates
         Return False

      End Function

      Private Overloads Function ParseCmd(ByVal asArgs() As String, Optional ByVal startIndex As Int32 = 0) As Boolean

         Dim sKey As String = String.Empty, sValue As String = String.Empty
         Dim bolResult As Boolean = True

         For i As Int32 = startIndex To asArgs.Length - 1
            bolResult = bolResult AndAlso ParseParam(asArgs(i))
         Next

         Return bolResult

      End Function

      Private Overloads Function ParseCmd(ByVal sArgs As String) As Boolean

         Dim asArgs As String() = sArgs.Split(CType(Me.DelimiterArgs, Char))
         Return ParseCmd(asArgs)

      End Function

      ''' <summary>
      ''' Parses a single key/pair combo into a matching <see cref="KeyValue"/> object
      ''' </summary>
      ''' <param name="sParam"></param>
      ''' <returns></returns>
      Private Function ParseParam(ByVal sParam As String) As Boolean

         If sParam.Length < 1 Then
            Return True
         End If

         With Me

            If sParam.Contains(.DelimiterValue) Then
               ' Parameter of the form /key=value

               Dim o As New KeyValue(sParam)

               With o
                  ' '/file' for /file=MyFile.txt
                  .KeyLong = Left(sParam, InStr(sParam, DelimiterValue) - 1).Trim
                  ' Remove the leading delimiter, results in 'file'
                  If .KeyLong.IndexOf(Me.DelimiterArgs) > -1 Then
                     .KeyLong = Mid(.KeyLong, Me.DelimiterArgs.Length + 1)
                  End If
                  ' Since we parse this from the command line, set both
                  .KeyShort = .KeyLong
                  .Value = Mid(sParam, InStr(sParam, DelimiterValue) + 1)
               End With

               .KeyValues.Add(o)

            Else
               ' Parameter of the form /Value.
               ' These are considered to be boolean parameters. If present, their value is 'True'

               Dim o As New KeyValue(sParam)

               With o
                  .KeyLong = sParam.Trim
                  If .KeyLong.IndexOf(Me.DelimiterArgs) > -1 Then
                     .KeyLong = Mid(.KeyLong, Me.DelimiterArgs.Length + 1)
                  End If
                  ' Since we parse this from the command line, set both
                  .KeyShort = .KeyLong
                  .Value = True
               End With

               .KeyValues.Add(o)

            End If

         End With

         Return True

      End Function


#End Region

#Region "Methods - Public"

      ''' <summary>
      ''' Initializes the object by parsing <see cref="System.Environment.GetCommandLineArgs"/>.
      ''' </summary>
      ''' <param name="cmdLineArgs">Parameters may optionally passed as a string.</param>
      ''' <returns><see langword="true"/> if parameters could successfully be parsed.</returns>
      Public Function Initialize(Optional ByVal cmdLineArgs As String = "") As Boolean

         ' Clear everything, as we're parsing anew.
         Me.KeyValues = New List(Of KeyValue)

         If cmdLineArgs.Length < 1 Then
            Dim asArgs() As String = System.Environment.GetCommandLineArgs()
            ' When using System.Environment.GetCommandLineArgs(), the 1st array element is the executable's name
            Return ParseCmd(asArgs, 1)
         Else
            Return ParseCmd(cmdLineArgs)
         End If

      End Function

      ''' <summary>
      ''' Validates all passed parameters
      ''' </summary>
      Public Sub Validate()

         ' Safe guard
         If Me.KeyValues.Count < 1 Then
            Exit Sub
         End If

         ' Any duplicates?
         For i As Int32 = 0 To Me.KeyValues.Count - 1
            Dim o As KeyValue = Me.KeyValues.Item(i)
            If HasDuplicate(o, i) = True Then
               Throw New ArgumentException(String.Format("Duplicate parameter: {0}", o.Key))
            End If
         Next

      End Sub

      ''' <summary>
      ''' Verify if a parameter was passed
      ''' </summary>
      ''' <param name="key">Name (<see cref="KeyValue.Key"/>) of the parameter</param>
      ''' <returns><see langword="true"/> if <paramref name="key"/> is present.</returns>
      ''' <remarks>If <see cref="CaseSensitive"/>=<see langword="true"/> then /param1 and /PARAM1 are 
      ''' treated as different parameters.</remarks>
      Public Overloads Function HasParameter(ByVal key As String) As Boolean

         With Me
            For Each o As KeyValue In .KeyValues
               If .CaseSensitive = True Then
                  If o.Key = key Then
                     Return True
                  End If
               Else
                  If o.Key.ToLower = key.ToLower Then
                     Return True
                  End If
               End If
            Next
         End With

         Return False

      End Function

      ''' <summary>
      ''' Determine if a certain parameter is present
      ''' </summary>
      ''' <param name="paramlist">List of parameter names (<see cref="KeyValue.Key"/>)</param>
      ''' <returns>
      ''' <see langword="true"/> if all parameters passed are present, otherwise <see langword="false"/>
      ''' </returns>
      Public Overloads Function HasParameter(ByVal paramList As List(Of String)) As Boolean

         With Me
            For Each s As String In paramList
               If HasParameter(s) = False Then
                  Return False
               End If
            Next
         End With

         ' Reaching here, all parameters have been found
         Return True

      End Function

      ''' <summary>
      ''' Determine if any of these parameters are present
      ''' </summary>
      ''' <param name="paramlist">List of parameter names (<see cref="KeyValue.Key"/>)</param>
      ''' <returns>
      ''' <see langword="true"/> if at least one of the parameters passed is present, otherwise <see langword="false"/>
      ''' </returns>
      Public Overloads Function HasAnyParameter(ByVal paramList As List(Of String)) As Boolean

         With Me
            For Each s As String In paramList
               If HasParameter(s) = True Then
                  Return True
               End If
            Next
         End With

         ' Reaching here, no parameter has been found
         Return False

      End Function

      ''' <summary>
      ''' Return the corresponding KeyValue object
      ''' </summary>
      ''' <param name="key">The parameter's name (<see cref="KeyValue.Key"/></param>
      ''' <param name="caseSensitive">Treat the name as case-sensitive?</param>
      ''' <returns><see cref="KeyValue"/> whose <see cref="KeyValue.Key"/> equals <paramref name="key"/>.</returns>
      Public Function GetParameterByName(ByVal key As String, Optional ByVal caseSensitive As Boolean = False) As KeyValue

         ' Safe guard
         If HasParameter(key) = False Then
            Throw New ArgumentException("Parameter doesn't exist: " & key)
         End If

         For Each o As KeyValue In Me.KeyValues
            If caseSensitive = False Then
               If o.Key.ToLower = key.ToLower Then
                  Return o
               End If
            Else
               If o.Key = key Then
                  Return o
               End If
            End If
         Next

         ' We should never reach this point
         Return Nothing

      End Function

      ''' <summary>
      ''' Return the value of a parameter
      ''' </summary>
      ''' <param name="key">The parameter's name (<see cref="KeyValue.Key"/></param>
      ''' <param name="caseSensitive">Treat the name as case-sensitive?</param>
      ''' <returns>Returns the <see cref="KeyValue.Value"/> of a parameter by its (parameter) name.</returns>
      Public Function GetValueByName(ByVal key As String, Optional ByVal caseSensitive As Boolean = False) As Object

         ' Safe guard
         If HasParameter(key) = False Then
            Throw New ArgumentException("Parameter doesn't exist: " & key)
         End If

         For Each o As KeyValue In Me.KeyValues
            If caseSensitive = False Then
               If o.Key.ToLower = key.ToLower Then
                  Return o.Value
               End If
            Else
               If o.Key = key Then
                  Return o.Value
               End If
            End If
         Next

         ' We should never reach this point
         Return Nothing

      End Function

#End Region

#Region "Constructor/Dispose"

      ''' <summary>
      ''' Object constructor.
      ''' </summary>
      Public Sub New()

         MyBase.New

         With Me
            .DelimiterArgs = GetDefaultDelimiterForOS()
            .DelimiterValue = DELIMITER_VALUE
            .ValidParameters = New List(Of String)
         End With

      End Sub

      ''' <summary>
      ''' Object constructor.
      ''' </summary>
      ''' <param name="delimiterArgs">Character which separates different parameters, e.g. "/" in /Param1=Value1 /Param2=Value2</param>
      ''' <param name="delimiterValue">Character which separates the parameter name and its value, e.g. "=" in /Param1=Value1 /Param2=Value2</param>
      Public Sub New(Optional ByVal delimiterArgs As String = DELIMITER_ARGS_WIN, Optional ByVal delimiterValue As String = DELIMITER_VALUE)

         MyBase.New

         ' Safe guard
         If delimiterArgs.Length < 1 OrElse delimiterValue.Length < 1 Then
            Throw New ArgumentOutOfRangeException("Empty argument or key/value delimiter are not allowed.")
         End If

         With Me
            .DelimiterArgs = delimiterArgs
            .DelimiterValue = delimiterValue
            .KeyValues = New List(Of KeyValue)
         End With

      End Sub

      ''' <summary>
      ''' Object constructor.
      ''' </summary>
      ''' <param name="delimiterArgsType">Character which separates different parameters as defined in <see cref="eArgumentDelimiterStyle"/></param>
      ''' <param name="delimiterValue">Character which separates the parameter name and its value, e.g. "=" in /Param1=Value1 /Param2=Value2</param>
      Public Sub New(Optional ByVal delimiterArgsType As eArgumentDelimiterStyle = eArgumentDelimiterStyle.Windows, Optional ByVal delimiterValue As String = DELIMITER_VALUE)

         MyBase.New

         ' Safe guard
         If delimiterValue.Length < 1 Then
            Throw New ArgumentOutOfRangeException("Empty argument or key/value delimiter are not allowed.")
         End If

         With Me
            If delimiterArgsType = eArgumentDelimiterStyle.Windows Then
               .DelimiterArgs = DELIMITER_ARGS_WIN
            Else
               .DelimiterArgs = DELIMITER_ARGS_POSIX
            End If
            .DelimiterValue = delimiterValue
            .KeyValues = New List(Of KeyValue)
         End With

      End Sub

      ''' <summary>
      ''' Object constructor.
      ''' </summary>
      ''' <param name="keyValueList">Command line parameters as a List(Of <see cref="KeyValue"/></param>
      ''' <param name="delimiterArgs">Character which separates different parameters, e.g. "/" in /Param1=Value1 /Param2=Value2</param>
      ''' <param name="delimiterValue">Character which separates the parameter name and its value, e.g. "=" in /Param1=Value1 /Param2=Value2</param>
      Public Sub New(ByVal keyValueList As List(Of KeyValue), Optional ByVal delimiterArgs As String = DELIMITER_ARGS_WIN,
                  Optional ByVal delimiterValue As String = DELIMITER_VALUE)

         MyBase.New

         ' Safe guard
         If delimiterArgs.Length < 1 OrElse delimiterValue.Length < 1 Then
            Throw New ArgumentOutOfRangeException("Empty argument or key/value delimiter are not allowed.")
         End If

         With Me
            .DelimiterArgs = delimiterArgs
            .DelimiterValue = delimiterValue
            .KeyValues = keyValueList
         End With

      End Sub


      ''' <summary>
      ''' Object constructor.
      ''' </summary>
      ''' <param name="keyValueList">Command line parameters as a List(Of <see cref="KeyValue"/></param>
      ''' <param name="delimiterArgsType">Character which separates different parameters as defined in <see cref="eArgumentDelimiterStyle"/></param>
      ''' <param name="delimiterValue">Character which separates the parameter name and its value, e.g. "=" in /Param1=Value1 /Param2=Value2</param>
      Public Sub New(ByVal keyValueList As List(Of KeyValue), Optional ByVal delimiterArgsType As eArgumentDelimiterStyle = eArgumentDelimiterStyle.Windows,
                  Optional ByVal delimiterValue As String = DELIMITER_VALUE)

         MyBase.New

         ' Safe guard
         If delimiterValue.Length < 1 Then
            Throw New ArgumentOutOfRangeException("Empty argument or key/value delimiter are not allowed.")
         End If

         With Me
            If delimiterArgsType = eArgumentDelimiterStyle.Windows Then
               .DelimiterArgs = DELIMITER_ARGS_WIN
            Else
               .DelimiterArgs = DELIMITER_ARGS_POSIX
            End If
            .DelimiterValue = delimiterValue
            .KeyValues = keyValueList
         End With

      End Sub

      Public Sub New(Optional ByVal validParams As List(Of String) = Nothing)

         MyBase.New

         With Me
            .DelimiterArgs = GetDefaultDelimiterForOS()
            .DelimiterValue = DELIMITER_VALUE
            If Not validParams Is Nothing Then
               .ValidParameters = validParams
            Else
               .ValidParameters = New List(Of String)
            End If
         End With

      End Sub

      Public Overloads Sub Dispose() Implements IDisposable.Dispose
         GC.SuppressFinalize(Me)
      End Sub

#End Region

   End Class

   ''' <summary>
   ''' A single command line parameter + its value, i.e. /parameter=value.
   ''' </summary>
   Public Class KeyValue
      Inherits KeyValueBase

      Private msHelpText As String = String.Empty
      Private msOriginalParameter As String = String.Empty

      Private moValue As Object

      ''' <summary>
      ''' Help text for this parameter
      ''' </summary>
      Public Property HelpText As String
         Get
            Return msHelpText
         End Get
         Set(value As String)
            msHelpText = value
         End Set
      End Property

      ''' <summary>
      ''' The full original parameter, e.g. /file=MyFile.txt
      ''' </summary>
      ''' <returns></returns>
      Public Property OriginalParameter As String
         Get
            Return msOriginalParameter
         End Get
         Set(value As String)
            msOriginalParameter = value
         End Set
      End Property

      ''' <summary>
      ''' Parameter name, e.g. 'file' or 'f' from /file=MyFile.txt  or /f=MyFile.txt
      ''' </summary>
      Public ReadOnly Property Key As String
         Get
            ' KeyLong takes precedences
            With Me
               If .KeyLong.Length > 0 Then
                  Return .KeyLong
               Else
                  Return .KeyShort
               End If
            End With
         End Get
      End Property

      ''' <summary>
      ''' Parameter value, e.g. 'MyFile.txt' from /file=MyFile.txt
      ''' </summary>
      Public Property Value As Object
         Get
            Return moValue
         End Get
         Set(value As Object)
            moValue = value
         End Set
      End Property

#Region "Methods - Public"

      ''' <summary>
      ''' Returns this parameter name.
      ''' </summary>
      ''' <returns><see cref="KeyValue.Key"/> and <see cref="KeyValue.HelpText"/> when available.</returns>
      Public Overrides Function ToString() As String

         With Me
            Dim sText As String = .Key
            If .HelpText.Length < 1 Then
               Return sText
            Else
               Return sText & ": " & .HelpText
            End If
         End With

      End Function

#End Region

      ''' <summary>
      ''' Create a new instance of this object.
      ''' </summary>
      ''' <param name="originalParameter">Key/value pair as originally passed via command line, e.g. /file=MyFile.txt</param>
      ''' <param name="keyShort">Short parameter name, e.g. /v</param>
      ''' <param name="keyLong">Verbose parameter name, e.g. /version</param>
      ''' <param name="value">Value of this parameter</param>
      ''' <param name="helpText">Explanatory text for this parameter</param>
      Public Sub New(ByVal originalParameter As String,
                     Optional ByVal keyShort As String = "",
                     Optional ByVal keyLong As String = "",
                     Optional ByVal value As Object = Nothing,
                     Optional ByVal helpText As String = "")

         MyBase.New

         With Me
            .HelpText = helpText
            .KeyLong = keyLong
            .KeyShort = keyShort
            .OriginalParameter = originalParameter
            .Value = value
         End With

      End Sub

   End Class

   ''' <summary>
   ''' Base/parent KeyValue class
   ''' </summary>
   Public Class KeyValueBase

      Private msKeyLong As String = String.Empty
      Private msKeyShort As String = String.Empty
      Private mbolIsMandatory As Boolean

      ''' <summary>
      ''' This parameter is mandatory
      ''' </summary>
      ''' <returns></returns>
      Public Property IsMandatory As Boolean
         Get
            Return mbolIsMandatory
         End Get
         Set(value As Boolean)
            mbolIsMandatory = value
         End Set
      End Property

      ''' <summary>
      ''' 'Outspoken' parameter name, e.g. /file
      ''' </summary>
      Public Property KeyLong As String
         Get
            Return msKeyLong
         End Get
         Set(value As String)
            msKeyLong = value
         End Set
      End Property

      ''' <summary>
      ''' Short parameter name, e.g. /f
      ''' </summary>
      Public Property KeyShort As String
         Get
            Return msKeyShort
         End Get
         Set(value As String)
            msKeyShort = value
         End Set
      End Property

      ''' <summary>
      ''' Short And long parameter name
      ''' <param name="shortKey">Short parameter name, e.g. /f</param>
      ''' <param name="longKey">'Outspoken' parameter name, e.g. /file</param>
      ''' </summary>
      Public Sub New(Optional ByVal shortKey As String = "", Optional ByVal longKey As String = "")
         KeyLong = longKey
         KeyShort = shortKey
      End Sub

      ''' <summary>
      ''' Short And long parameter name
      ''' </summary>
      ''' <param name="o">
      ''' <see cref="KeyValueBase"/> object from which to take <see cref="KeyValueBase.KeyShort"/> and <see cref="KeyValue.KeyLong"/> names.
      ''' </param>
      Public Sub New(ByVal o As KeyValueBase)
         KeyLong = o.KeyLong
         KeyShort = o.KeyShort
      End Sub

   End Class

End Namespace
