Public Module GlobalVariables
    'Speed Unit
    Public mpsTOkmph As Decimal = 3.6
    Public mpsTOmlph As Decimal = 2.236936
    Public mpsTOfpm As Decimal = 196.850394
    Public mlpsTOkmph As Decimal = 5793.6384
    Public mlpsTOmlph As Decimal = 3600
    Public mlpsTOfpm As Decimal = 316800
    Public fpsTOkmps As Decimal = 1.09728
    Public fpsTOmlph As Decimal = 0.681818
    Public fpsTOfpm As Decimal = 60

    Public kmphTOmps As Decimal = 0.277778
    Public mlphTOmps As Decimal = 0.44704
    Public fpmTOmps As Decimal = 0.00508
    Public kmphTOmlps As Decimal = 0.000173
    Public mlphTOmlps As Decimal = 0.000278
    Public fpmTOmlps As Decimal = 0.000003
    Public kmphTOfps As Decimal = 0.911344
    Public mlphTOfps As Decimal = 1.466667
    Public fpmTOfps As Decimal = 0.016667

    'Energy Unit
    Public joulesTOfcalories As Decimal = 0.000239
    Public joulesTOeV As Decimal = 6241457005723417000
    Public joulesTONm As Decimal = 1
    Public caloriesTOfcalories As Decimal = 0.001
    Public caloriesTOeV As Decimal = 2.6131952590564E+19
    Public caloriesTONm As Decimal = 4.1868
    Public kwhTOfcalories As Decimal = 859845.227859
    Public kwhTOeV As Decimal = 2.24692452206043E+25
    Public kwhTONm As Decimal = 3600000

    Public fcaloriesTOjoules As Decimal = 4186
    Public eVTOjoules As Decimal = 1.602176487E-19
    Public NmTOjoules As Decimal = 1
    Public fcaloriesTOcalories As Decimal = 999.808923
    Public eVTOcalories As Decimal = 3.8267327959301E-20
    Public NmTOcalories As Decimal = 0.238846
    Public fcaloriesTOkwh As Decimal = 0.000001
    Public eVTOkwh As Decimal = 4.4504902416667E-26
    Public NmTOkwh As Decimal = 0.000000278

    'Mass Unit
    Public gTOkg As Decimal = 0.001
    Public gTOounces As Decimal = 0.035274
    Public gTOslugs As Decimal = 0.000069
    Public poundsTOkg As Decimal = 0.453592
    Public poundsTOounces As Decimal = 16
    Public poundsTOslugs As Decimal = 0.031081
    Public tonsTOkg As Decimal = 1016.046909
    Public tonsTOounces As Decimal = 35840
    Public tonsTOslugs As Decimal = 69.621328

    Public kgTOg As Decimal = 1000
    Public ouncesTOg As Decimal = 28.349523
    Public slugsTOg As Decimal = 14593.903
    Public kgTOpounds As Decimal = 2.204623
    Public ouncesTOpounds As Decimal = 0.0625
    Public slugsTOpounds As Decimal = 32.174049
    Public kgTOtons As Decimal = 0.000984
    Public ouncesTOtons As Decimal = 0.000028
    Public slugsTOtons As Decimal = 0.014363

End Module
Public Class Form1
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox1.ReadOnly = True
    End Sub
    Public Sub SpeedConverter()
        Dim StartingUnit As String = ComboBox1.Text
        Dim DestinationUnit As String = ComboBox2.Text
        Dim input As Decimal
        Dim result As Decimal

        'convert string to integer
        input = CDec(TextBox2.Text)

        'Calculation
        Select Case StartingUnit
            Case "meters per second"
                Select Case DestinationUnit
                    Case "kilometers per hour"
                        result = input * mpsTOkmph
                    Case "miles per hour"
                        result = input * mpsTOmlph
                    Case "feet per minute"
                        result = input * mpsTOfpm
                End Select
            Case "miles per second"
                Select Case DestinationUnit
                    Case "kilometers per hour"
                        result = input * mlpsTOkmph
                    Case "miles per hour"
                        result = input * mlpsTOmlph
                    Case "feet per minute"
                        result = input * mlpsTOfpm
                End Select
            Case "feet per second"
                Select Case DestinationUnit
                    Case "kilometers per hour"
                        result = input * fpsTOkmps
                    Case "miles per hour"
                        result = input * fpsTOmlph
                    Case "feet per minute"
                        result = input * fpsTOfpm
                End Select
            Case "kilometers per hour"
                Select Case DestinationUnit
                    Case "meters per second"
                        result = input * kmphTOmps
                    Case "miles per second"
                        result = input * kmphTOmlps
                    Case "feet per second"
                        result = input * kmphTOfps
                End Select
            Case "miles per hour"
                Select Case DestinationUnit
                    Case "meters per second"
                        result = input * mlphTOmps
                    Case "miles per second"
                        result = input * mlphTOmlps
                    Case "feet per second"
                        result = input * mlphTOfps
                End Select
            Case "feet per minute"
                Select Case DestinationUnit
                    Case "meters per second"
                        result = input * fpmTOmps
                    Case "miles per second"
                        result = input * fpmTOmlps
                    Case "feet per second"
                        result = input * fpmTOfps
                End Select
        End Select
        TextBox1.Text = result

    End Sub
    Public Sub EnergyConverter()
        Dim StartingUnit As String = ComboBox1.Text
        Dim DestinationUnit As String = ComboBox2.Text
        Dim input As Decimal
        Dim result As Decimal

        'convert string to integer
        input = CDec(TextBox2.Text)

        'Calculation
        Select Case StartingUnit
            Case "joules"
                Select Case DestinationUnit
                    Case "calories(food)"
                        result = input * joulesTOfcalories
                    Case "electron volts"
                        result = input * joulesTOeV
                    Case "newton meters"
                        result = input * joulesTONm
                End Select
            Case "calories"
                Select Case DestinationUnit
                    Case "calories(food)"
                        result = input * caloriesTOfcalories
                    Case "electron volts"
                        result = input * caloriesTOeV
                    Case "newton meters"
                        result = input * caloriesTONm
                End Select
            Case "kilowatt hours"
                Select Case DestinationUnit
                    Case "calories(food)"
                        result = input * kwhTOfcalories
                    Case "electron volts"
                        result = input * kwhTOeV
                    Case "newton meters"
                        result = input * kwhTONm
                End Select
            Case "calories(food)"
                Select Case DestinationUnit
                    Case "joules"
                        result = input * fcaloriesTOjoules
                    Case "calories"
                        result = input * fcaloriesTOcalories
                    Case "kilowatt hours"
                        result = input * fcaloriesTOkwh
                End Select
            Case "electron volts"
                Select Case DestinationUnit
                    Case "joules"
                        result = input * eVTOjoules
                    Case "calories"
                        result = input * eVTOcalories
                    Case "kilowatt hours"
                        result = input * eVTOkwh
                End Select
            Case "newton meters"
                Select Case DestinationUnit
                    Case "joules"
                        result = input * NmTOjoules
                    Case "calories"
                        result = input * NmTOcalories
                    Case "kilowatt hours"
                        result = input * NmTOkwh
                End Select
        End Select
        TextBox1.Text = result

    End Sub
    Public Sub MassConverter()
        Dim StartingUnit As String = ComboBox1.Text
        Dim DestinationUnit As String = ComboBox2.Text
        Dim input As Decimal
        Dim result As Decimal

        'convert string to integer
        input = CDec(TextBox2.Text)

        'Calculation
        Select Case StartingUnit
            Case "grams"
                Select Case DestinationUnit
                    Case "kilograms"
                        result = input * gTOkg
                    Case "ounces"
                        result = input * gTOounces
                    Case "slugs"
                        result = input * gTOslugs
                End Select
            Case "pounds"
                Select Case DestinationUnit
                    Case "kilograms"
                        result = input * poundsTOkg
                    Case "ounces"
                        result = input * poundsTOounces
                    Case "slugs"
                        result = input * poundsTOslugs
                End Select
            Case "tons"
                Select Case DestinationUnit
                    Case "kilograms"
                        result = input * tonsTOkg
                    Case "ounces"
                        result = input * tonsTOounces
                    Case "slugs"
                        result = input * tonsTOslugs
                End Select
            Case "kilograms"
                Select Case DestinationUnit
                    Case "grams"
                        result = input * kgTOg
                    Case "pounds"
                        result = input * kgTOpounds
                    Case "tons"
                        result = input * kgTOtons
                End Select
            Case "ounces"
                Select Case DestinationUnit
                    Case "grams"
                        result = input * ouncesTOg
                    Case "pounds"
                        result = input * ouncesTOpounds
                    Case "tons"
                        result = input * ouncesTOtons
                End Select
            Case "slugs"
                Select Case DestinationUnit
                    Case "grams"
                        result = input * slugsTOg
                    Case "pounds"
                        result = input * slugsTOpounds
                    Case "tons"
                        result = input * slugsTOtons
                End Select
        End Select
        TextBox1.Text = result

    End Sub
    Private Sub executepart()
        Dim Number As Boolean = True

        If TextBox2.Text = "" Then
            TextBox1.Text = ""
        Else
            Number = check()
            If Number = False Then
                MessageBox.Show("Invalid Character, Please enter a number")
            Else
                Select Case ComboBox3.Text()
                    Case "Speed"
                        SpeedConverter()
                    Case "Energy"
                        EnergyConverter()
                    Case "Mass"
                        MassConverter()
                End Select
            End If
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        executepart()
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        executepart()
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        executepart()
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim SpeedUnit1() As String = {"meters per second", "miles per second", "feet per second"}
        Dim SpeedUnit2() As String = {"kilometers per hour", "miles per hour", "feet per minute"}
        Dim EnergyUnit1() As String = {"joules", "calories", "kilowatt hours"}
        Dim EnergyUnit2() As String = {"calories(food)", "electron volts", "newton meters"}
        Dim MassUnit1() As String = {"grams", "pounds", "tons"}
        Dim MassUnit2() As String = {"kilograms", "ounces", "slugs"}

        If ComboBox3.Text = "" Then
            ComboBox1.Items.Clear()
            ComboBox2.Items.Clear()
        Else
            Select Case ComboBox3.SelectedItem
                Case "Speed"
                    TextBox1.Clear()
                    TextBox2.Clear()
                    ComboBox1.Items.Clear()
                    ComboBox2.Items.Clear()
                    ComboBox1.Items.AddRange(SpeedUnit1)
                    ComboBox2.Items.AddRange(SpeedUnit2)
                Case "Energy"
                    TextBox1.Clear()
                    TextBox2.Clear()
                    ComboBox1.Items.Clear()
                    ComboBox2.Items.Clear()
                    ComboBox1.Items.AddRange(EnergyUnit1)
                    ComboBox2.Items.AddRange(EnergyUnit2)
                Case "Mass"
                    TextBox1.Clear()
                    TextBox2.Clear()
                    ComboBox1.Items.Clear()
                    ComboBox2.Items.Clear()
                    ComboBox1.Items.AddRange(MassUnit1)
                    ComboBox2.Items.AddRange(MassUnit2)
            End Select
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Num1 As Integer = ComboBox1.Items.Count - 1
        Dim Num2 As Integer = ComboBox2.Items.Count - 1
        Dim Items1(Num1) As String
        Dim items2(Num2) As String
        Dim i As Integer
        Dim j As Integer
        Dim CurrentItem1 As String = ComboBox1.SelectedItem
        Dim CurrentItem2 As String = ComboBox2.SelectedItem

        'Get all items from combobox1
        For i = 0 To Num1
            Items1(i) = ComboBox1.Items(i)
        Next

        'Get all items from combobox2
        For j = 0 To Num2
            items2(j) = ComboBox2.Items(j)
        Next

        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()

        For i = 0 To Num1
            ComboBox2.Items.Add(Items1(i))
        Next

        For j = 0 To Num2
            ComboBox1.Items.Add(items2(j))
        Next

        ComboBox1.Text = CurrentItem2
        ComboBox2.Text = CurrentItem1

    End Sub
    Public Function check() As Boolean
        Dim i As Integer = TextBox2.Text.Count - 1
        Dim NumArray() As Char
        Dim value As Integer
        Dim correct As Boolean = True

        NumArray = TextBox2.Text.ToCharArray

        For i = 0 To i
            value = Asc(NumArray(i))
            If value < 47 Or value > 58 Then
                correct = False
            End If

            If value = 46 Then
                correct = True
            End If
        Next

        Return correct

    End Function
End Class
