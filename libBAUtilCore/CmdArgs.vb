'------------------------------------------------------------------------------
'   Author: Knuth Konrad 2021-04-05
'  Changed: -
'------------------------------------------------------------------------------
Imports libBAUtilCore.StringUtil

Namespace Utils.CmdArgs

   ''' <summary>
   ''' Commandline parameter handling
   ''' See https://docs.microsoft.com/en-us/archive/msdn-magazine/2019/march/net-parse-the-command-line-with-system-commandline
   ''' See https://github.com/commandlineparser/commandline
   ''' </summary>
   Public Class CmdArgs

#Region "Declarations"

      Public Enum eArgumentDelimiterStyle
         Windows
         POSIX
      End Enum


      ' Default arguments and key/value delimiter
      Private Const DELIMITER_ARGS_WIN As String = "/"
      Private Const DELIMITER_ARGS_POSIX As String = "--"
      Private Const DELIMITER_VALUE As String = "="

      Dim mbolCaseSensitive As Boolean                ' Treat parameter names as case-sensitive?
      Dim msDelimiterArgs As String = String.Empty    ' Arguments delimiter, typically "/"
      Dim msDelimiterValue As String = String.Empty   ' Key/value delimiter, typically "="
      Dim listValidParameters As List(Of String)      ' List of all valid parameter

      Private mcolKeyValues As List(Of KeyValue)

#End Region

#Region "Properties - Public"

      Public ReadOnly Property ParametersCount As Int32
         Get
            If Not Me.KeyValues Is Nothing Then
               Return Me.KeyValues.Count
            Else
               Return 0
            End If
         End Get
      End Property

      Public Property CaseSensitive As Boolean
         Get
            Return mbolCaseSensitive
         End Get
         Set(value As Boolean)
            mbolCaseSensitive = value
         End Set
      End Property

      Public Property DelimiterArgs As String
         Get
            Return msDelimiterArgs
         End Get
         Set(value As String)
            msDelimiterArgs = value
         End Set
      End Property

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
                  ElseIf (o.Key = currentKey.Key) AndAlso (i <> currentIndex) Then
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
      ''' Initializes the object by parsing System.Environment.GetCommandLineArgs()
      ''' </summary>
      ''' <returns></returns>
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
      ''' Determine if a certain parameter is present
      ''' </summary>
      ''' <param name="key">Parameter's name (<see cref="KeyValue.Key"/>)</param>
      ''' <returns></returns>
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
               If o.Key.ToLower = key Then
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
      ''' <returns></returns>
      Public Function GetValueByName(ByVal key As String, Optional ByVal caseSensitive As Boolean = False) As Object

         ' Safe guard
         If HasParameter(key) = False Then
            Throw New ArgumentException("Parameter doesn't exist: " & key)
         End If

         For Each o As KeyValue In Me.KeyValues
            If caseSensitive = False Then
               If o.Key.ToLower = key Then
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


      Public Sub New()

         MyBase.New

         With Me
            .DelimiterArgs = DELIMITER_ARGS_WIN
            .DelimiterValue = DELIMITER_VALUE
            .ValidParameters = New List(Of String)
         End With

      End Sub

      Public Sub New(Optional ByVal validParams As List(Of String) = Nothing)

         MyBase.New

         With Me
            .DelimiterArgs = DELIMITER_ARGS_WIN
            .DelimiterValue = DELIMITER_VALUE
            If Not validParams Is Nothing Then
               .ValidParameters = validParams
            Else
               .ValidParameters = New List(Of String)
            End If
         End With

      End Sub

      Public Sub New(Optional ByVal delimiterArgs As String = DELIMITER_ARGS_WIN, Optional ByVal delimiterValue As String = DELIMITER_VALUE,
                     Optional ByVal validParams As List(Of String) = Nothing)

         MyBase.New

         ' Safe guard
         If delimiterArgs.Length < 1 OrElse delimiterValue.Length < 1 Then
            Throw New ArgumentOutOfRangeException("Empty argument or key/value delimiter are not allowed.")
         End If

         With Me
            .DelimiterArgs = delimiterArgs
            .DelimiterValue = delimiterValue
            .KeyValues = New List(Of KeyValue)
            If Not validParams Is Nothing Then
               .ValidParameters = validParams
            Else
               .ValidParameters = New List(Of String)
            End If
         End With

      End Sub

      Public Sub New(Optional ByVal delimiterArgsType As eArgumentDelimiterStyle = eArgumentDelimiterStyle.Windows,
                     Optional ByVal delimiterValue As String = DELIMITER_VALUE,
                     Optional ByVal validParams As List(Of String) = Nothing)

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
            If Not validParams Is Nothing Then
               .ValidParameters = validParams
            Else
               .ValidParameters = New List(Of String)
            End If
         End With

      End Sub

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
            ' Create the list of valid parameter names from the passed collection
            .ValidParameters = New List(Of String)
            For Each o As KeyValue In .KeyValues
               .ValidParameters.Add(o.Key)
            Next
         End With

      End Sub

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
            ' Create the list of valid parameter names from the passed collection
            .ValidParameters = New List(Of String)
            For Each o As KeyValue In .KeyValues
               .ValidParameters.Add(o.Key)
            Next
         End With

      End Sub

#End Region

   End Class

   Public Class KeyValue

      Private msHelpText As String = String.Empty
      Private msKeyLong As String = String.Empty
      Private msKeyShort As String = String.Empty
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

      Public Sub New(ByVal originalParam As String, Optional ByVal keyShort As String = "",
                     Optional ByVal keyLong As String = "", Optional ByVal value As Object = Nothing,
                     Optional ByVal helpText As String = "")

         MyBase.New

         With Me
            .HelpText = helpText
            .KeyLong = keyLong
            .KeyShort = keyShort
            .OriginalParameter = originalParam
            .Value = value
         End With

      End Sub

   End Class

End Namespace

