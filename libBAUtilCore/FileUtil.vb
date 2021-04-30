Imports System.IO
Imports System.Text

''' <summary>
''' Simple text file manipulation.
''' </summary>
Public Class TextFileUtil

   ''' <summary>
   ''' Read a text file from disk and return it as a string
   ''' </summary>
   ''' <param name="textFile">Text file incl. full path</param>
   ''' <returns>Contents of <paramref name="textFile"/> as a string</returns>
   Public Shared Function TxtReadFile(ByVal textFile As String) As String

      Dim file As String = String.Empty

      Using reader As New StreamReader(textFile)
         file = reader.ReadToEnd
      End Using

      Return file

   End Function

   ''' <summary>
   ''' Retrieve a line of a text file and return it as a string
   ''' </summary>
   ''' <param name="textFile">Text file incl. full path</param>
   ''' <returns>The line at the current position of the text file</returns>
   Public Shared Function TxtReadLine(ByVal textFile As String) As String

      Dim line As String = String.Empty

      Using reader As New StreamReader(textFile)
         line = reader.ReadLine()
      End Using

      Return line

   End Function

   ''' <summary>
   ''' Write a line to a text file
   ''' </summary>
   ''' <param name="textFile">Text file incl. full path</param>
   ''' <param name="textLine">Contents of new line</param>
   ''' <param name="doAppend">
   ''' <see langref="true"/>: append <paramref name="textLine"/> at the end of <paramref name="textFile"/><br />
   ''' <see langref="false"/>: insert <paramref name="textLine"/> at the current position of <paramref name="textFile"/>
   ''' </param>
   Public Shared Sub TxtWriteLine(ByVal textFile As String, ByVal textLine As String, Optional ByVal doAppend As Boolean = True)

      Using writer As New StreamWriter(textFile, doAppend)
         writer.WriteLine(textLine)
      End Using

   End Sub

End Class
