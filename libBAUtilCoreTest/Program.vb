Imports System
Imports libBAUtilCore.ConsoleHelper


Module Program
  Sub Main(args As String())
    Console.WriteLine("Hello World!")

    ConsoleHide()
    System.Threading.Thread.Sleep(3000)
    ConsoleShow()

    AnyKey()

  End Sub
End Module
