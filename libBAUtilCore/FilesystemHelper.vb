Imports System.IO
Imports System.IO.File
Imports System.IO.Path

''' <summary>
''' General file system helper methods
''' </summary>
Public Class FilesystemHelper


   ''' <summary>
   ''' Standard Windows path delimiter.
   ''' </summary>
   Private Const DELIMITER_PATH_WIN As String = "\"
   ''' <summary>
   ''' Standard POSIX ("Linux") path delimiter.
   ''' </summary>
   Private Const DELIMITER_PATH_POSIX As String = "/"

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
   Public Shared Function DenormalizePath(ByVal sPath As String, ByVal sDelim As String,
                                          Optional bolCheckTail As Boolean = True) As String
      If bolCheckTail = True Then
         If StringHelper.Right(sPath, sDelim.Length) <> sDelim Then
            Return sPath
         Else
            Return sPath.Substring(1, sPath.Length - sDelim.Length)
         End If
      Else
         If StringHelper.Left(sPath, sDelim.Length) <> sDelim Then
            Return sPath
         Else
            Return sPath.Substring(sDelim.Length + 1)
         End If
      End If
   End Function

   ''' <summary>
   ''' Ensure a path does NOT end with a path delimiter
   ''' </summary>
   ''' <param name="sPath">
   ''' Path (drive or network share), C:\Windows\ or \\myserver\myshare\
   ''' </param>
   ''' <param name="bolCheckTail">
   ''' Check the start or end (default) of <paramref name="sPath"/>.
   ''' </param>
   ''' <returns>
   ''' <paramref name="sPath"/> with <paramref name="sDelim"/> stripped off, if present.
   ''' </returns>
   Public Shared Function DenormalizePath(ByVal sPath As String,
                                          Optional bolCheckTail As Boolean = True) As String
      Dim sDelim As String = Path.PathSeparator

      If bolCheckTail = True Then
         If StringHelper.Right(sPath, sDelim.Length) <> sDelim Then
            Return sPath
         Else
            Return sPath.Substring(1, sPath.Length - sDelim.Length)
         End If
      Else
         If StringHelper.Left(sPath, sDelim.Length) <> sDelim Then
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
   Public Shared Function NormalizePath(ByVal sPath As String, ByVal sDelim As String,
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
         If StringHelper.Right(sPath, sDelim.Length) <> sDelim Then
            Return sPath & sDelim
         Else
            Return sPath
         End If
      Else
         If StringHelper.Left(sPath, sDelim.Length) <> sDelim Then
            Return sDelim & sPath
         Else
            Return sPath
         End If
      End If

   End Function

   ''' <summary>
   ''' Ensure a path does end with a path delimiter
   ''' </summary>
   ''' <param name="sPath">
   ''' Path (drive or network share), C:\Windows\ or \\myserver\myshare\
   ''' </param>
   ''' <param name="bolCheckTail">
   ''' Check the start or end (default) of <paramref name="sPath"/>.
   ''' </param>
   ''' <returns>
   ''' <paramref name="sPath"/> with OS.specific path separator add, if not present.
   ''' </returns>
   Public Shared Function NormalizePath(ByVal sPath As String,
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
      Dim sDelim As String = Path.PathSeparator

      If bolCheckTail = True Then
         If StringHelper.Right(sPath, sDelim.Length) <> sDelim Then
            Return sPath & sDelim
         Else
            Return sPath
         End If
      Else
         If StringHelper.Left(sPath, sDelim.Length) <> sDelim Then
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
   ''' Alias for <see cref="System.IO.Directory.Exists(String)"/>
   ''' </summary>
   ''' <param name="folder">The file to check.</param>
   ''' <returns><see langword="true"/> if the caller has the required permissions and path contains the name of an existing file; otherwise, <see langword="false"/>. 
   ''' This method also returns <see langword="false"/> if path is <see langword="null"/>, an invalid path, or a zero-length string. 
   ''' If the caller does not have sufficient permissions to read the specified file, no exception is thrown and the method returns 
   ''' <see langword="false"/> regardless of the existence of path.
   ''' </returns>
   Public Shared Function FolderExists(ByVal folder As String) As Boolean
      Return Directory.Exists(folder)
   End Function


   ''' <summary>
   ''' Retrieve the parameter delimiter according to the OS' typical flavor
   ''' </summary>
   ''' <returns>OS typical parameter delimiter</returns>
   Public Shared Function GetDefaultPathDelimiterForOS() As String
      If System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows) Then
         Return DELIMITER_PATH_WIN
      Else
         Return DELIMITER_PATH_POSIX
      End If
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


   Public Shared Function ShortenPathText(ByVal sPath As String, ByVal lMaxLen As Int32, Optional ByVal sDelim As String = "") As String
      '------------------------------------------------------------------------------
      'Funktion : Kürzt eine Pfadangabe auf lMaxLen Zeichen
      '
      'Vorauss. : -
      'Parameter: sPath    -  zu kürzende Pfadangabe
      '           lMaxLen  -  maximal Länge des Pfades
      'Rückgabe : -
      '
      'Autor    : Doberenz & Kowalski
      'erstellt : 26.11.1999
      'geändert : Knuth Konrad
      '           ungarische Notation und Stringfunktion statt Variant (Mid, Left...)
      'Notiz    : Quelle: Visual Basic 6 Kochbuch, Hanser Verlag
      '------------------------------------------------------------------------------
      Dim i, lLen, lDiff As Int32
      Dim sTemp As String = String.Empty

      lLen = sPath.Length

      If lLen <= lMaxLen Then
         Return sPath
      End If

      If sDelim.Length < 1 Then
         sDelim = Path.PathSeparator
      End If

      For i = (lLen - lMaxLen + 6) To lLen
         If StringHelper.Mid(sPath, i, 1) = sDelim Then Exit For
      Next i

      If StringHelper.InStr(sPath, sDelim) < 1 Then
         ' Ist wohl nur eine Datei, ohne Pfadangabe -> die "Mitte" des Namens kürzen

         sTemp = sPath

         If lLen > lMaxLen Then
            lDiff = lLen - lMaxLen
         Else
            lDiff = 0
         End If

         lDiff = lDiff \ 2

         If lDiff > 2 Then
            sTemp = StringHelper.Left(sPath, lLen \ 2) & "..." & StringHelper.Right(sPath, lLen \ 2)
         End If

      Else
         If i < lLen Then
            sTemp = StringHelper.Left(sPath, 3) & "..." & StringHelper.Right(sPath, lLen - (i - 1))
         Else
            sTemp = StringHelper.Left(sPath, 3) & "..." & StringHelper.Right(sPath, lMaxLen - 6)
         End If
      End If

      Return sTemp

   End Function

End Class
