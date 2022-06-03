Imports libBAUtilCore.StringHelper

''' <summary>
''' Network (files, folders, shares) related utilities.
''' Please note: it's Windows only!
''' </summary>
Public Class NetworkUtil

   Private Declare Ansi Function WNetGetConnection Lib "mpr.dll" Alias _
        "WNetGetConnectionA" _
          (ByVal lpszLocalName As String,
          ByVal lpszRemoteName As String,
          ByRef cbRemoteName As Int32) As Int32


   Public Shared Function UNCPathFromDriveLetter(ByVal sPath As String, ByRef dwError As Int32,
                                                 Optional ByVal returnDriveOnly As Boolean = False) As String
      '------------------------------------------------------------------------------
      'Purpose  : Returns a fully qualified UNC path location from a (mapped network)
      '           drive letter/share
      '
      'Prereq.  : -
      'Parameter: sPath       - Path to resolve
      '           dwError     - ByRef(!), Returns the error code from the Win32 API, if any
      '           lDriveOnly  - If True, return only the drive letter
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 17.07.2013
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
      ' 32-bit declarations:
      Dim sTemp As String = String.Empty
      Dim sDrive As String = String.Empty
      Dim lStatus As Int32

      ' ** Safe guards
      ' Prevent this class from being used on non-Windows OS
      If System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(Runtime.InteropServices.OSPlatform.Windows) Then
         Throw New System.ApplicationException("This object is only available on Windows platforms.")
      End If


      Const NO_ERROR As Int32 = 0

      ' The size used for the string buffer. Adjust this if you
      ' need a larger buffer.
      Dim lBUFFER_SIZE As Int32 = 1024
      Dim sRemoteName As String = Space(lBUFFER_SIZE)

      If sPath.Length > 2 Then
         sTemp = Mid(sPath, 3)
         sDrive = Left(sPath, 2)
      Else
         sDrive = sPath
      End If

      ' Return the UNC path (\\Server\Share).
      lStatus = WNetGetConnection(sDrive, sRemoteName, lBUFFER_SIZE)

      ' Verify that the WNetGetConnection() succeeded. WNetGetConnection()
      ' returns 0 (NO_ERROR) if it successfully retrieves the UNC path.
      If lStatus = NO_ERROR Then
         ' Display the UNC path.

         If returnDriveOnly = True Then
            UNCPathFromDriveLetter = TrimAny(sRemoteName, Chr(0) & vbWhiteSpace())
         Else
            UNCPathFromDriveLetter = TrimAny(sRemoteName, Chr(0) & vbWhiteSpace()) & sTemp
         End If

      Else
         ' Return the original filename/path unaltered
         UNCPathFromDriveLetter = sPath
      End If

      dwError = lStatus

   End Function

#Region "Constructor / Finalizer"

   Public Sub New()

      MyBase.New

      ' Prevent this class from being used on non-Windows OS
      If System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(Runtime.InteropServices.OSPlatform.Windows) Then
         Throw New System.ApplicationException("This object is only available on Windows platforms.")
      End If

   End Sub

#End Region

End Class
