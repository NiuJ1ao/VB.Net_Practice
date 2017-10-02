Module Module1
    Dim array(4) As String
    Dim addingls(4) As Integer
    Sub Main()
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim i As Integer

        x = Factorial(6)

        array(0) = "a"
        array(1) = "b"
        array(2) = "c"
        array(3) = "d"
        array(4) = "e"
        y = LinearSearch("d", 0)

        addingls(0) = 5
        addingls(1) = 4
        addingls(2) = 6
        addingls(3) = 10
        addingls(4) = 15
        z = Adding(0)

        i = Fibonacci(9)
        Console.WriteLine(x)
        Console.WriteLine(y)
        Console.WriteLine(z)
        Console.WriteLine(i)
        TimeTable(5, 1)
        Console.ReadLine()
    End Sub

    Public Function Factorial(ByVal n As Integer) As Integer
        Dim result As Integer

        If n = 1 Then
            result = 1
        Else
            result = n * Factorial(n - 1)
        End If

        Return result

    End Function

    Public Function LinearSearch(ByVal str As String, ByVal i As Integer) As Integer
        Dim found As Integer

        If i = 5 Then
            found = -1
        ElseIf str = array(i) Then
            found = i
        Else
            found = LinearSearch(str, i + 1)
        End If

        Return found
    End Function

    Public Function Adding(ByVal i As Integer) As Integer
        Dim r As Integer

        If i = 5 Then
            r = 0
        Else
            r = addingls(i) + Adding(i + 1)
        End If

        Return r
    End Function

    Public Function Fibonacci(ByVal n As Integer) As Integer
        Dim r As Integer

        If n <= 2 Then
            r = 1
        Else
            r = Fibonacci(n - 1) + Fibonacci(n - 2)
        End If

        Return r
    End Function

    Public Sub TimeTable(ByVal multiple As Integer, ByVal count As Integer)
        Console.WriteLine(multiple & " x " & count & " = " & multiple * count)

        If count > 11 Then

        Else
            TimeTable(multiple, count + 1)
        End If
    End Sub
End Module
