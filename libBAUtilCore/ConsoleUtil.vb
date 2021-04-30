Imports System.Reflection

Imports libBAUtilCore.StringUtil

''' <summary>
''' General purpose console application helpers
''' </summary>
Public Class ConsoleUtil

#Region "Declarations"
   ' Copyright notice values
   Private Const COPY_COMPANYNAME As String = "BasicAware"
   Private Const COPY_AUTHOR As String = "Knuth Konrad"
   ' Console defaults
   Private Const CON_SEPARATOR As String = "---"

#End Region

#Region "ConHeadline"

   ''' <summary>
   ''' Display an application intro
   ''' </summary>
   ''' <param name="appName">Name of the application</param>
   ''' <param name="versionMajor">Major version</param>
   ''' <param name="versionMinor">Minor version</param>
   ''' <param name="versionRevision">Revision</param>
   ''' <param name="versionBuild">Build</param>
   Public Overloads Shared Sub ConHeadline(ByVal appName As String, ByVal versionMajor As Integer,
                                           Optional ByVal versionMinor As Integer = 0,
                                           Optional ByVal versionRevision As Int32 = 0,
                                           Optional versionBuild As Int32 = 0)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine(Chr(16) & " " & appName & " v" &
                        versionMajor.ToString & "." &
                        versionMinor.ToString & "." &
                        versionRevision.ToString & "." &
                        versionBuild.ToString &
                        " " & Chr(17))
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   ''' <summary>
   ''' Display an application intro
   ''' </summary>
   ''' <param name="appName">Name of the application</param>
   Public Overloads Shared Sub ConHeadline(ByVal appName As String)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine(Chr(16) & " " & appName & " " & Chr(17))
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   ''' <summary>
   ''' Display an application intro
   ''' </summary>
   ''' <param name="appName">Name of the application</param>
   ''' <param name="versionMajor">Major version</param>
   Public Overloads Shared Sub ConHeadline(ByVal appName As String, ByVal versionMajor As Integer)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine(Chr(16) & " " & appName & " v" & versionMajor.ToString & ".0 " & Chr(17))
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   ''' <summary>
   ''' Display an application intro
   ''' </summary>
   ''' <param name="mainAssembly">Assembly of the main executable</param>
   ''' <example>
   ''' ConHeadline(System.Refelction.Assembly.GetEntryAssembly())
   ''' </example>
   ''' <remarks>
   ''' See https://docs.microsoft.com/en-us/dotnet/api/system.version?view=net-5.0
   ''' </remarks>
   Public Overloads Shared Sub ConHeadline(ByVal mainAssembly As Assembly)

      Dim assemName As AssemblyName = mainAssembly.GetName()
      Dim ver As Version = assemName.Version

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine(Chr(16) & " {0} v{1}", assemName.Name, ver.ToString() & " " & Chr(17))
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub
#End Region

#Region "ConCopyright"
   ''' <summary>
   ''' Display a copyright notice.
   ''' </summary>
   Public Overloads Shared Sub ConCopyright(Optional ByVal trailingBlankLine As Boolean = True)
      ConCopyright(DateTime.Now.Year.ToString, COPY_COMPANYNAME, trailingBlankLine)
   End Sub

   ''' <summary>
   ''' Display a copyright notice.
   ''' </summary>
   ''' <param name="companyName">Copyright owner</param>
   Public Overloads Shared Sub ConCopyright(ByVal companyName As String, Optional ByVal trailingBlankLine As Boolean = True)
      ConCopyright(DateTime.Now.Year.ToString, companyName, trailingBlankLine)
   End Sub

   ''' <summary>
   ''' Display a copyright notice.
   ''' </summary>
   ''' <param name="year">Copyrighted in year</param>
   ''' <param name="companyName">Copyright owner</param>
   Public Overloads Shared Sub ConCopyright(ByVal year As String, ByVal companyName As String, Optional ByVal trailingBlankLine As Boolean = True)
      Console.WriteLine(String.Format("Copyright {0} {1} by {2}. All rights reserved.", Chr(169), year, companyName))
      Console.WriteLine("Written by " & COPY_AUTHOR)

      If trailingBlankLine = True Then
         Console.WriteLine("")
      End If
   End Sub
#End Region

   ''' <summary>
   ''' Pauses the program execution and waits for a key press
   ''' </summary>
   ''' <param name="waitMessage">Pause message</param>
   ''' <param name="blankLinesBefore">Number of blank lines before the message</param>
   ''' <param name="blankLinesAfter">Number of blank lines after the message</param>
   Public Shared Sub AnyKey(Optional ByVal waitMessage As String = "-- Press ENTER to continue --",
                            Optional ByVal blankLinesBefore As Int32 = 0,
                            Optional ByVal blankLinesAfter As Int32 = 0)

      BlankLine(blankLinesBefore)
      Console.WriteLine(waitMessage)
      BlankLine(blankLinesAfter)
      Console.ReadLine()

   End Sub

   ''' <summary>
   ''' Insert a blank line at the current position.
   ''' </summary>
   ''' <param name="blankLines">Number of blank lines to insert.</param>
   ''' <param name="addSeparatingLine"><see langword="true"/>: Add a visual separation indicator before the blank line(s)</param>
   Public Overloads Shared Sub BlankLine(Optional ByVal blankLines As Int32 = 1, Optional ByVal addSeparatingLine As Boolean = False)

      ' Safe guard
      If blankLines < 1 Then
         blankLines = 1
      End If

      If addSeparatingLine = True Then
         Console.WriteLine(CON_SEPARATOR)
      End If

      For i As Int32 = 0 To blankLines - 1
         Console.WriteLine("")
      Next

   End Sub

   ''' <summary>
   ''' Output text indented by (<paramref name="indentBy"/>) spaces
   ''' </summary>
   ''' <param name="text">Output text</param>
   ''' <param name="indentBy">Number of leading spaces</param>
   ''' <param name="addNewLine">Add a new line after <paramref name="text"/></param>
   Public Shared Sub WriteIndent(ByVal text As String, ByVal indentBy As Int32, Optional ByVal addNewLine As Boolean = True)

      If addNewLine = True Then
         Console.WriteLine(String.Concat(New String(CType(" ", Char), indentBy) & text))
      Else
         Console.Write(String.Concat(New String(CType(" ", Char), indentBy) & text))
      End If

   End Sub

End Class
