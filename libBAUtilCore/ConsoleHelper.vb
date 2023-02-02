Imports System.Reflection
Imports System.Runtime.InteropServices

Imports libBAUtilCore.StringHelper
Imports libBAUtilCore.ConHelperData

''' <summary>
''' General purpose console application helpers
''' </summary>
Public Class ConsoleHelper

#Region "Declarations"
  ' Copyright notice values
  ' Private Const COPY_COMPANYNAME As String = "BasicAware"
  ' Private Const COPY_AUTHOR As String = "Knuth Konrad"
  ' Console defaults
  ' Private Const CON_SEPARATOR As String = "---"

  ' Console cursor position WaitIndicator
  Private mThdCurLeft, mThdCurTop, mThdSpinDelay As Int32
  Private mThdNewLine As Boolean
  Private thdWaitIndicator As System.Threading.Thread

  ' Show/hide console window
  Const SW_HIDE As Integer = 0
  Const SW_SHOW As Integer = 5
  Shared handle As IntPtr = GetConsoleWindow()
  <DllImport("kernel32.dll")>
  Shared Function GetConsoleWindow() As IntPtr
  End Function
  <DllImport("user32.dll")>
  Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
  End Function

  Private cts As Threading.CancellationTokenSource = New Threading.CancellationTokenSource
#End Region

#Region "AppIntro"

  ''' <summary>
  ''' Display an application intro
  ''' </summary>
  ''' <param name="appName">Name of the application</param>
  ''' <param name="versionMajor">Major version</param>
  ''' <param name="versionMinor">Minor version</param>
  ''' <param name="versionRevision">Revision</param>
  ''' <param name="versionBuild">Build</param>
  Public Overloads Shared Sub AppIntro(ByVal appName As String, ByVal versionMajor As Integer,
                                           Optional ByVal versionMinor As Integer = 0,
                                           Optional ByVal versionRevision As Int32 = 0,
                                           Optional versionBuild As Int32 = 0)

    Console.ForegroundColor = ConsoleColor.White
    Console.WriteLine("* " & appName & " v" &
                      versionMajor.ToString & "." &
                      versionMinor.ToString & "." &
                      versionRevision.ToString & "." &
                      versionBuild.ToString &
                      " *")
    Console.ForegroundColor = ConsoleColor.Gray

  End Sub

  ''' <summary>
  ''' Display an application intro
  ''' </summary>
  ''' <param name="appName">Name of the application</param>
  Public Overloads Shared Sub AppIntro(ByVal appName As String)

    Console.ForegroundColor = ConsoleColor.White
    Console.WriteLine("* " & appName & " *")
    Console.ForegroundColor = ConsoleColor.Gray

  End Sub

  ''' <summary>
  ''' Display an application intro
  ''' </summary>
  ''' <param name="appName">Name of the application</param>
  ''' <param name="versionMajor">Major version</param>
  Public Overloads Shared Sub AppIntro(ByVal appName As String, ByVal versionMajor As Integer)

    Console.ForegroundColor = ConsoleColor.White
    Console.WriteLine("* " & appName & " v" & versionMajor.ToString & ".0 *")
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
  Public Overloads Shared Sub AppIntro(ByVal mainAssembly As Assembly)

    Dim assemName As AssemblyName = mainAssembly.GetName()
    Dim ver As Version = assemName.Version

    Console.ForegroundColor = ConsoleColor.White
    Console.WriteLine(Chr(16) & " {0} v{1}", assemName.Name, ver.ToString() & " " & Chr(17))
    Console.ForegroundColor = ConsoleColor.Gray

  End Sub


#End Region

#Region "AppCopyright"

  ''' <summary>
  ''' Display a copyright notice.
  ''' </summary>
  ''' <param name="trailingBlankLine">Add a blank line afterwards.</param>
  Public Overloads Shared Sub AppCopyright(Optional ByVal trailingBlankLine As Boolean = True)
    AppCopyright(DateTime.Now.Year.ToString, ConHelperData.COPY_COMPANYNAME, trailingBlankLine)
  End Sub

  ''' <summary>
  ''' Display a copyright notice.
  ''' </summary>
  ''' <param name="companyName">Copyright owner</param>
  ''' <param name="trailingBlankLine">Add a blank line afterwards.</param>
  Public Overloads Shared Sub AppCopyright(ByVal companyName As String, Optional ByVal trailingBlankLine As Boolean = True)
    AppCopyright(DateTime.Now.Year.ToString, companyName, trailingBlankLine)
  End Sub

  ''' <summary>
  ''' Display a copyright notice.
  ''' </summary>
  ''' <param name="year">Copyrighted in year</param>
  ''' <param name="companyName">Copyright owner</param>
  ''' <param name="trailingBlankLine">Add a blank line afterwards.</param>
  Public Overloads Shared Sub AppCopyright(ByVal year As String, ByVal companyName As String, Optional ByVal trailingBlankLine As Boolean = True)
    Console.WriteLine(String.Format("Copyright {0} {1} by {2}. All rights reserved.", Chr(169), year, companyName))
    Console.WriteLine("Written by " & ConHelperData.COPY_AUTHOR)

    If trailingBlankLine = True Then
      Console.WriteLine("")
    End If
  End Sub

#End Region

#Region "AnyKey"
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

#End Region

#Region "BlankLine"
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
#End Region

#Region "ConsoleHide"
  ''' <summary>
  ''' Hides the console window
  ''' </summary>
  ''' <remarks>Show/Hide source: https://superuser.com/questions/398605/how-to-force-windows-desktop-background-to-update-or-refresh</remarks>
  Public Shared Sub ConsoleHide()
    ' Hide the console
    ShowWindow(handle, SW_HIDE)
  End Sub
#End Region

  ''' <summary>
  ''' Shows the console windows
  ''' </summary>
#Region "ConsoleShow"
  Public Shared Sub ConsoleShow()
    ' Show the console
    ShowWindow(handle, SW_SHOW)
  End Sub
#End Region

#Region "WriteIndent"
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
  ''' <summary>
  ''' Output text indented by (<paramref name="indentBy"/>) spaces
  ''' </summary>
  ''' <param name="text">Output text</param>
  ''' <param name="indentBy">Number of leading spaces</param>
  ''' <param name="addNewLine">Add a new line after the last line of <paramref name="text"/></param>
  Public Shared Sub WriteIndent(text As String(), indentBy As Int32, Optional addNewLine As Boolean = True)
    Dim i As Integer = 0
    While i <= text.Length - 2
      Console.WriteLine(String.Concat(New [String](Convert.ToChar(" "), indentBy) + text(i)))
      i += 1
    End While
    If addNewLine = True Then
      Console.WriteLine(String.Concat(New [String](Convert.ToChar(" "), indentBy) + text(text.Length - 1)))
    Else
      Console.Write(String.Concat(New [String](Convert.ToChar(" "), indentBy) + text(text.Length - 1)))
    End If
  End Sub
#End Region

#Region "WriteOK"

  ''' <summary>
  ''' Output text in the 'OK' color
  ''' </summary>
  ''' <param name="text">Output text</param>
  ''' <param name="bgColor">Background color.</param>
  ''' <param name="fgColor">Foreground color.</param>
  ''' <param name="addNewLine">Add a new line after the last line of <paramref name="text"/></param>
  Public Shared Sub WriteOK(text As String, Optional fgColor As ConsoleColor = ConsoleColor.Green, Optional bgColor As ConsoleColor = ConsoleColor.Black, Optional addNewLine As Boolean = True)

    Console.BackgroundColor = bgColor
    Console.ForegroundColor = fgColor
    WriteIndent(text, 0, addNewLine)
    Console.ResetColor()

  End Sub

  ''' <summary>
  ''' Output text indented by (<paramref name="indentBy"/>) spaces in the 'OK' color
  ''' </summary>
  ''' <param name="text">Output text</param>
  ''' <param name="indentBy">Number of leading spaces</param>
  ''' <param name="bgColor">Background color.</param>
  ''' <param name="fgColor">Foreground color.</param>
  ''' <param name="addNewLine">Add a new line after the last line of <paramref name="text"/></param>
  Public Shared Sub WriteOK(text As String, indentBy As Int32, Optional fgColor As ConsoleColor = ConsoleColor.Green, Optional bgColor As ConsoleColor = ConsoleColor.Black, Optional addNewLine As Boolean = True)

    Console.BackgroundColor = bgColor
    Console.ForegroundColor = fgColor
    WriteIndent(text, indentBy, addNewLine)
    Console.ResetColor()

  End Sub

  ''' <summary>
  ''' Output text indented by (<paramref name="indentBy"/>) spaces in the 'OK' color
  ''' </summary>
  ''' <param name="text">Output text</param>
  ''' <param name="indentBy">Number of leading spaces</param>
  ''' <param name="bgColor">Background color.</param>
  ''' <param name="fgColor">Foreground color.</param>
  ''' <param name="addNewLine">Add a new line after the last line of <paramref name="text"/></param>
  Public Shared Sub WriteOK(text As String(), indentBy As Int32, Optional fgColor As ConsoleColor = ConsoleColor.Green, Optional bgColor As ConsoleColor = ConsoleColor.Black, Optional addNewLine As Boolean = True)

    Console.BackgroundColor = bgColor
    Console.ForegroundColor = fgColor
    WriteIndent(text, indentBy, addNewLine)
    Console.ResetColor()

  End Sub

#End Region

#Region "WriteError"

  ''' <summary>
  ''' Output text in the 'Error' color
  ''' </summary>
  ''' <param name="text">Output text</param>
  ''' <param name="bgColor">Background color.</param>
  ''' <param name="fgColor">Foreground color.</param>
  ''' <param name="addNewLine">Add a new line after the last line of <paramref name="text"/></param>
  Public Shared Sub WriteError(text As String, Optional fgColor As ConsoleColor = ConsoleColor.Red, Optional bgColor As ConsoleColor = ConsoleColor.Black, Optional addNewLine As Boolean = True)

    Console.BackgroundColor = bgColor
    Console.ForegroundColor = fgColor
    WriteIndent(text, 0, addNewLine)
    Console.ResetColor()

  End Sub

  ''' <summary>
  ''' Output text indented by (<paramref name="indentBy"/>) spaces in the 'Error' color
  ''' </summary>
  ''' <param name="text">Output text</param>
  ''' <param name="indentBy">Number of leading spaces</param>
  ''' <param name="bgColor">Background color.</param>
  ''' <param name="fgColor">Foreground color.</param>
  ''' <param name="addNewLine">Add a new line after the last line of <paramref name="text"/></param>
  Public Shared Sub WriteError(text As String, indentBy As Int32, Optional fgColor As ConsoleColor = ConsoleColor.Red, Optional bgColor As ConsoleColor = ConsoleColor.Black, Optional addNewLine As Boolean = True)

    Console.BackgroundColor = bgColor
    Console.ForegroundColor = fgColor
    WriteIndent(text, indentBy, addNewLine)
    Console.ResetColor()

  End Sub

  ''' <summary>
  ''' Output text indented by (<paramref name="indentBy"/>) spaces in the 'Error' color
  ''' </summary>
  ''' <param name="text">Output text</param>
  ''' <param name="indentBy">Number of leading spaces</param>
  ''' <param name="bgColor">Background color.</param>
  ''' <param name="fgColor">Foreground color.</param>
  ''' <param name="addNewLine">Add a new line after the last line of <paramref name="text"/></param>
  Public Shared Sub WriteError(text As String(), indentBy As Int32, Optional fgColor As ConsoleColor = ConsoleColor.Red, Optional bgColor As ConsoleColor = ConsoleColor.Black, Optional addNewLine As Boolean = True)

    Console.BackgroundColor = bgColor
    Console.ForegroundColor = fgColor
    WriteIndent(text, indentBy, addNewLine)
    Console.ResetColor()

  End Sub
#End Region

#Region "WaitIndicator"

  ''' <summary>
  ''' Starts a "spinning wheel" kinda wait time indicator
  ''' </summary>
  ''' <param name="newLine">Add a newline on the first post?</param>
  ''' <param name="spinDelay">A "spinning tick" occurs every <paramref name="spinDelay"/> milliseconds.</param>
  Public Sub WaitIndicatorStart(Optional ByVal newLine As Boolean = True, Optional ByVal spinDelay As Int32 = 100)

    mThdCurLeft = Console.CursorLeft
    mThdCurTop = Console.CursorTop
    mThdNewLine = newLine
    mThdSpinDelay = spinDelay
    thdWaitIndicator = New Threading.Thread(AddressOf WaitIndicator)
    thdWaitIndicator.Start()

  End Sub

  ''' <summary>
  ''' Stops a previously started <see cref="WaitIndicatorStart(Boolean, Integer)"/> wait time indicator.
  ''' </summary>
  Public Sub WaitIndicatorStop()

    If Not thdWaitIndicator Is Nothing Then
      ' thdWaitIndicator.Abort()
      cts.Cancel()
      thdWaitIndicator.Join()
      cts.Dispose()
    End If

  End Sub

  ''' <summary>
  ''' Object Finalizer
  ''' </summary>
  Protected Overrides Sub Finalize()
    WaitIndicatorStop()
    MyBase.Finalize()
  End Sub

  Private Sub WaitIndicator()

    Dim lStep As Int32
    Dim blnBeenHere As Boolean = False
    Dim curLeftOld, curTopOld As Int32, curVisibleOld As Boolean

    Dim asChar() As String = {"/", "-", "\", "|"}

    ' Safe guard
    If Console.IsOutputRedirected = True Then
      Exit Sub
    End If

    Do
      Threading.Thread.BeginCriticalRegion()

      curLeftOld = Console.CursorLeft
      curTopOld = Console.CursorTop
      If System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows) Then
        curVisibleOld = Console.CursorVisible
      End If

      Console.SetCursorPosition(mThdCurLeft, mThdCurTop)
      Console.CursorVisible = False
      Console.Write(asChar(lStep))

      If blnBeenHere = False Then
        If mThdNewLine = True Then
          curTopOld += 1
          curLeftOld = 0
        End If
        blnBeenHere = True
      End If

      lStep += 1
      If lStep >= 3 Then
        lStep = 0
      End If
      Console.SetCursorPosition(curLeftOld, curTopOld)
      Console.CursorVisible = curVisibleOld
      Threading.Thread.EndCriticalRegion()

      Threading.Thread.Sleep(mThdSpinDelay)

      If cts.IsCancellationRequested = True Then
        Exit Do
      End If
    Loop

  End Sub

#End Region

End Class
