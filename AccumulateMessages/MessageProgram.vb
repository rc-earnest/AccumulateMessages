'Rudy Earnest
'RCET 2265
'Spring 2025
'Accumulitive Messages
'https://github.com/rc-earnest/AccumulateMessages.git   
Option Strict On
Option Explicit On
Option Compare Text

Imports System

Module MessageProgram
    Sub Main(args As String())
        'uncomment to test interactively
        'Test.Manual()
        Test.Auto()
    End Sub

    Function UserMessages(ByVal newMessage As String, ByVal clear As Boolean) As String
        Static message As New Text.StringBuilder() 'dims message as a "global variable" and makes it as a string builder class

        If clear Then
            message.Clear() 'clears the message string builder

        ElseIf Not String.IsNullOrEmpty(newMessage) Then 'checks to make sure that there is something inside of the string
            message.AppendLine(newMessage) 'writes every string in a newline that is in newMessage to message

        End If
        Return message.ToString() 'converts the message string builder to string
    End Function



End Module
