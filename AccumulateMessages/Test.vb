
Option Explicit On
Option Strict On
Option Compare Text
Class Test
    ''' <summary>
    ''' Manually verify function
    ''' </summary>
    Public Shared Sub Manual()

        Dim userInput As String
        Do

            Select Case userInput
                Case "D"
                    Console.Clear()
                    Console.WriteLine("Messages:" & vbCrLf)
                    Console.WriteLine(UserMessages("", False))
                Case "C"
                    Console.Clear()
                    Console.WriteLine("Messages Cleared")
                    UserMessages(userInput, True)
                Case Else
                    UserMessages(userInput, False)
                    userInput = ""
                    Console.Clear()
                    Console.WriteLine("Please type a message:" & vbCrLf & vbCrLf _
                              & " D: Display all saved messages" & vbCrLf _
                              & " C: Clear all saved messages" & vbCrLf _
                              & " Q: Quit"
                              )
            End Select

            Console.WriteLine(vbCrLf & "Press Enter to continue")
            userInput = Console.ReadLine()
        Loop While userInput <> "Q"

        Console.Clear()
        Console.WriteLine("Have a nice day!")

    End Sub
    ''' <summary>
    ''' Automatic testing. Make all tests pass!
    ''' </summary>
    Public Shared Sub Auto()
        Dim expected$, actual$, clearAfter%
        Dim testdata() = {"Hello",
                          "Good bye",
                          "Jimmy likes pizza!!",
                          "too many bananas",
                          "more",
                          "aardvark",
                          "must be a number",
                          "I need one more message"
                          }
        Console.WriteLine("Call with empty strings:")
        expected = ""
        For i = 0 To 2
            UserMessages("", False)
        Next

        AreEqual(expected, UserMessages("", False))

        Console.WriteLine("Call with many sequential messages:")
        expected = ""
        For i = LBound(testdata) To UBound(testdata)
            expected += testdata(i) & vbCrLf
            actual = UserMessages(testdata(i), False)
        Next

        AreEqual(expected, actual)

        Console.WriteLine("Call with clear messages True:")
        expected = ""
        actual = UserMessages("", True)
        AreEqual(expected, actual)
        actual = UserMessages("Anything", True)
        AreEqual(expected, actual)

        Console.WriteLine("Call with clear midway in sequential messages:")
        expected = ""
        clearAfter = RandomNumberInRange(7, 3)
        For i = LBound(testdata) To UBound(testdata)
            If i = clearAfter Then
                expected = ""
                UserMessages("", True)
            End If

            expected += testdata(i) & vbCrLf
            actual = UserMessages(testdata(i), False)
        Next

        AreEqual(expected, actual)

    End Sub

    Private Shared Function AreEqual(expected$, Actual$) As Boolean
        Dim textcolor As ConsoleColor
        Dim result$

        Select Case Actual
            Case expected
                textcolor = ConsoleColor.Green
                result = "PASS"
            Case Else
                textcolor = ConsoleColor.Red
                result = "FAIL"
        End Select

        Console.ForegroundColor = textcolor
        Console.WriteLine($"{StrDup(4, "*")} {result}: {StrDup(4, "*")}")
        Console.WriteLine($">> Expected:")
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine(expected$)
        Console.ForegroundColor = textcolor
        Console.WriteLine($">> Actual:")
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine(Actual)
        Console.ForegroundColor = textcolor
        Console.WriteLine()
        'Console.WriteLine(StrDup(15, "*"))
        Console.ForegroundColor = ConsoleColor.White

    End Function

    Private Shared Function RandomNumberInRange(Optional max% = 10%, Optional min% = 0%) As Integer
        Dim _max% = max - min
        Randomize(DateTime.Now.Millisecond)
        Return CInt(System.Math.Floor(Rnd() * (_max + 1))) + min
    End Function
End Class


