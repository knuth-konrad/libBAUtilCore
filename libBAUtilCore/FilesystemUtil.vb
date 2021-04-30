Imports System.IO
Imports System.IO.File
Imports System.IO.Path

''' <summary>
''' General file system helper methods
''' </summary>
Public Class FilesystemUtil

   ''' <summary>
   ''' Ensure a path does NOT end with a path delimiter
   ''' </summary>
   ''' <param name="sPath">
   ''' Path (drive or network share), C:\Windows\ or \\myserver\myshare\
   ''' </param>
   ''' <param name="sDelim">
   ''' Character to be treated as the folder delimiter.
   ''' </param>
   ''' <param name="bolCheckTail">
   ''' Check the start or end (default) of <paramref name="sPath"/>.
   ''' </param>
   ''' <returns>
   ''' <paramref name="sPath"/> with <paramref name="sDelim"/> stripped off, if present.
   ''' </returns>
   Public Shared Function DenormalizePath(ByVal sPath As String, Optional ByVal sDelim As String = "\",
                                          Optional bolCheckTail As Boolean = True) As String
      '------------------------------------------------------------------------------
      'Prereq.  : -
      'Note     : -
      '
      '   Author: Bruce McKinney - Hardcore Visual Basic 5
      '     Date: 08.04.2019
      '   Source: -
      '  Changed: 09.08.1999, Knuth Konrad
      '           - Pass argument(s) ByVal and String instead of Variant
      '------------------------------------------------------------------------------
      If bolCheckTail = True Then
         If StringUtil.Right(sPath, sDelim.Length) <> sDelim Then
            Return sPath
         Else
            Return sPath.Substring(1, sPath.Length - sDelim.Length)
         End If
      Else
         If StringUtil.Left(sPath, sDelim.Length) <> sDelim Then
            Return sPath
         Else
            Return sPath.Substring(sDelim.Length + 1)
         End If
      End If
   End Function

   ''' <summary>
   ''' Ensure a path does end with a path delimiter
   ''' </summary>
   ''' <param name="sPath">
   ''' Path (drive or network share), C:\Windows\ or \\myserver\myshare\
   ''' </param>
   ''' <param name="sDelim">
   ''' Character to be treated as the folder delimiter.
   ''' </param>
   ''' <param name="bolCheckTail">
   ''' Check the start or end (default) of <paramref name="sPath"/>.
   ''' </param>
   ''' <returns>
   ''' <paramref name="sPath"/> with <paramref name="sDelim"/> add, if not present.
   ''' </returns>
   Public Shared Function NormalizePath(ByVal sPath As String, Optional ByVal sDelim As String = "\",
                                        Optional bolCheckTail As Boolean = True) As String
      '------------------------------------------------------------------------------
      'Prereq.  : -
      '
      '   Author: Bruce McKinney - Hardcore Visual Basic 5
      '     Date: 07.06.2018
      '   Source: -
      '   Change: 09.08.1999, Knuth Konrad
      '           - Pass argument(s) ByVal and String instead of Variant
      '------------------------------------------------------------------------------
      If bolCheckTail = True Then
         If StringUtil.Right(sPath, sDelim.Length) <> sDelim Then
            Return sPath & sDelim
         Else
            Return sPath
         End If
      Else
         If StringUtil.Left(sPath, sDelim.Length) <> sDelim Then
            Return sDelim & sPath
         Else
            Return sPath
         End If
      End If

   End Function

   ''' <summary>
   ''' Alias for <see cref="System.IO.File.Exists(String)"/>
   ''' </summary>
   ''' <param name="file">The file to check.</param>
   ''' <returns><see langword="true"/> if the caller has the required permissions and path contains the name of an existing file; otherwise, <see langword="false"/>. 
   ''' This method also returns <see langword="false"/> if path is <see langword="null"/>, an invalid path, or a zero-length string. 
   ''' If the caller does not have sufficient permissions to read the specified file, no exception is thrown and the method returns 
   ''' <see langword="false"/> regardless of the existence of path.
   ''' </returns>
   Public Shared Function FileExists(ByVal [file] As String) As Boolean
      Return Exists(file)
   End Function

   ''' <summary>
   ''' Creates a backup of a file by copying/moving it from the source 
   ''' to the target folder.
   ''' </summary>
   ''' <param name="fileSource">
   ''' Fully qualified source file name.
   ''' </param>
   ''' <param name="fileDest">
   ''' Fully qualified destination file name.
   ''' </param>
   ''' <param name="copyOnly">
   ''' Copy only (True) or move (False)?
   ''' </param>
   ''' <param name="incrementTarget">
   ''' If <paramref name="fileDest"/> exists, create a new file named 
   ''' (file).(nnnn).(ext) instead?
   ''' Please note: duplicate creation is limited to 10,000 files.
   ''' </param>
   ''' <param name="newFile">
   ''' (ByRef!) Returns the (newly created) fully qualified destination file name.
   ''' </param>
   ''' <returns>
   ''' Success: <see langword="true"/>, failure: <see langword="false"/>.
   ''' </returns>
   ''' <remarks>
   ''' If the destination file already exists, but <paramref name="incrementTarget"/> = <see langword="false"/>, 
   ''' the already existing file will be overwritten.
   ''' </remarks>
   Public Shared Function BackupFile(ByVal fileSource As String, ByVal fileDest As String,
                              Optional ByVal copyOnly As Boolean = False,
                              Optional ByVal incrementTarget As Boolean = True,
                              Optional ByRef newFile As String = "") As Boolean
      '------------------------------------------------------------------------------
      'Prereq.  : -
      'Note     : 
      '
      '           
      '
      '   Author: Knuth Konrad 03.05.2019
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim tempFile As String = String.Empty

      Dim destPath As String = String.Empty
      Dim destFile As String = String.Empty
      Dim destExt As String = String.Empty

      ' Safe guard
      If FileExists(fileSource) = False Then
         Return False
      End If

      ' Check if the destination file already exists
      ' Split the file name into its components path, file name, extension
      ' Create a new backup file name (incrementTarget = True)
      If FileExists(fileDest) Then

         ' Split the file name into its components
         destPath = GetDirectoryName(fileDest)
         destFile = GetFileNameWithoutExtension(fileDest)
         destExt = GetExtension(fileDest)

         ' Target already exists, create a copy instead?
         Dim i As Int32
         If incrementTarget = True Then
            i = 0
            ' Generate a new file name that eventually doesn't exist in the target folder. Stop at 9999!
            tempFile = NormalizePath(destPath) & destFile & "." & String.Format("{0:0000}", i) & destExt
            Do While Exists(tempFile) = True And i < 9999
               i += 1
               tempFile = NormalizePath(destPath) & destFile & "." & String.Format("{0:0000}", i) & destExt
            Loop
         Else
            tempFile = fileDest
         End If
      Else
         tempFile = fileDest
      End If

      If copyOnly = True Then
         File.Copy(fileSource, tempFile)
      Else
         File.Move(fileSource, tempFile)
      End If

      newFile = tempFile

      Return True

   End Function

End Class
