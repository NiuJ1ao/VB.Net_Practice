Module Module1

    Sub Main()
        Dim choice As Char
        Dim UserNote As New Note

        Do
            Console.Clear()
            Console.WriteLine("+------------------------+")
            Console.WriteLine("|A. Create new class     |")
            Console.WriteLine("|B. File already exist?  |")
            Console.WriteLine("|Q. QUIT                 |")
            Console.WriteLine("+------------------------+")
            Console.Write("What is your choice: ")
            choice = Console.ReadLine()

            Select Case choice
                Case "A"
                    UserNote.SetNewFile()
                Case "B"
                    UserNote.GetFile()
            End Select

        Loop Until choice = "Q"


    End Sub

End Module
