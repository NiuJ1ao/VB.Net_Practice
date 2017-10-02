Public Class Note
    Private ClassName As String
    Private ClassID As Integer
    Private Defn As Defination

    Public Sub New()

    End Sub

    Public Sub SetNewFile()
        Dim NewName As String

        Console.Write("Class name:")
        NewName = Console.ReadLine()

        NewName = NewName & ".txt"
        Dim NewWriter As New IO.StreamWriter(NewName)
        NewWriter.WriteLine("Format of record:")
        NewWriter.WriteLine("Keyword: Defination")
        NewWriter.Close()

        Defn = New Defination
        Defn.AddNew(NewName)

    End Sub

    Public Sub GetFile()
        Dim choice As Char
        Dim UserFile As String
        Dim UserWord As String

        Console.WriteLine("+-----------------------+")
        Console.WriteLine("| A. Add new defination |")
        Console.WriteLine("| B. Search defination  |")
        Console.WriteLine("+-----------------------+")
        Console.Write("Please input your choice: ")
        choice = Console.ReadLine()

        Defn = New Defination
        Select Case choice
            Case "A"
                Console.Write("Which class/file are you looking for? Please input file name: ")
                UserFile = Console.ReadLine()
                UserFile = UserFile

                Defn.AddNew(UserFile)

            Case "B"
                Console.Write("Which class/file are you looking for? Please input file name: ")
                UserFile = Console.ReadLine()

                Console.Write("Which word are you looking for? ")
                UserWord = Console.ReadLine()
                Defn.GetDefn(UserFile, UserWord)
        End Select
    End Sub

End Class
