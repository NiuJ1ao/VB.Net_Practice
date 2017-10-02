Public Class Defination

    Public Sub New()

    End Sub
    Structure Defination
        Dim Kword As String
        Dim ex As String
    End Structure

    Public Sub AddNew(ByVal FileName As String)
        Dim AddWriter As New IO.StreamWriter(FileName, True)
        Dim NewDefination() As Defination
        Dim i As Integer

        Console.Write("How many definations do you want to add: ")
        i = Console.ReadLine()

        ReDim NewDefination(i)

        For i = 1 To i
            Console.Write("Keyword: ")
            NewDefination(i).Kword = Console.ReadLine()
            Console.Write("Defination: ")
            NewDefination(i).ex = Console.ReadLine()

            AddWriter.WriteLine(NewDefination(i).Kword & ": " & NewDefination(i).ex)
        Next

        AddWriter.Close()
        Console.WriteLine("The new record have been saved!")

    End Sub

    Public Sub GetDefn(ByVal UserFile As String, ByVal UserWord As String)
        Dim FileName As String = (UserFile & ".txt")
        Dim FileReader As New IO.StreamReader(FileName)
        Dim Found As Boolean = False
        Dim CurrentLine As String = ""
        Dim CurrentLineArray() As Char
        Dim isColon As Boolean = False
        Dim i As Integer
        Dim CurrentP As Integer
        Dim Word As String

        Do While FileReader.Peek() <> -1 And Found = False
            CurrentLine = FileReader.ReadLine()
            CurrentLineArray = CurrentLine.ToCharArray
            For i = 0 To CurrentLine.Length() - 1
                If CurrentLineArray(i) = ":" Then
                    CurrentP = i
                    Word = Left(CurrentLine, CurrentP)
                    If Word = UserWord Then
                        Found = True
                    End If
                End If
            Next
        Loop

        FileReader.Close()
        Console.WriteLine(CurrentLine)
        Console.ReadLine()
    End Sub

End Class
