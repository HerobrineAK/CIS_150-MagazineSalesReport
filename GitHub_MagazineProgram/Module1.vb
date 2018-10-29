Module Module1
    '   Current Record to Process
    Private CurrentRecord() As String

    '   Sales Person Info
    Private SalesPersonNumber As String
    Private SalesPersonName As String

    '   Magazine Sales Variables
    Private LIFE_Sales As Integer
    Private TIME_Sales As Integer
    Private USNEWS_Sales As Integer
    Private TotalMagazineSales As Integer

    '   Bonus Pay Calculation
    Private BonusPay As Decimal
    Private Const BonusPayRate As Decimal = 0.05

    '   Info File
    Private MagazineFile As New Microsoft.VisualBasic.FileIO.TextFieldParser("MAGS.txt")

    Sub Main()
        Call HouseKeeping()
        Do While Not (MagazineFile).EndOfData
            Call ProcessRecords()
        Loop
        Call EndOfJob()
    End Sub

    Sub HouseKeeping()
        Call SetFileDelimiters()
        Call WriteHeadings()
    End Sub

    Sub SetFileDelimiters()
        MagazineFile.TextFieldType = FileIO.FieldType.Delimited
        MagazineFile.SetDelimiters(",")
    End Sub

    Sub WriteHeadings()
        Console.WriteLine()
        Console.WriteLine(Space(30) & "Magazine Sales Report")
        Console.WriteLine(Space(33) & "By Salespersons")
        Console.WriteLine()
        Console.WriteLine(Space(2) & "Salesperson" & Space(7) & "Life" & Space(6) & "Time" & Space(5) & "US News" & Space(8) & "Total" & Space(9) & "Bonus")
        Console.WriteLine(Space(2) & "Name" & Space(13) & "Sales" & Space(5) & "Sales" & Space(7) & "Sales" & Space(8) & "Sales" & Space(11) & "Pay")
        Console.WriteLine()
    End Sub
    Sub ProcessRecords()
        Call ReadFile()
        Call DetailCalculation()
        Call WriteDetailLine()
    End Sub

    Sub ReadFile()
        CurrentRecord = MagazineFile.ReadFields()

        SalesPersonName = CurrentRecord(1)
        SalesPersonNumber = CurrentRecord(0)

        LIFE_Sales = CurrentRecord(2)
        TIME_Sales = CurrentRecord(3)
        USNEWS_Sales = CurrentRecord(4)
    End Sub

    Sub DetailCalculation()
        TotalMagazineSales = LIFE_Sales + TIME_Sales + USNEWS_Sales
        BonusPay = TotalMagazineSales * BonusPayRate
    End Sub

    Sub WriteDetailLine()
        Console.WriteLine(Space(2) & SalesPersonName.ToString().PadRight(10) & Space(9) & LIFE_Sales.ToString().PadLeft(3) & Space(7) & TIME_Sales.ToString().PadLeft(3) & Space(9) & USNEWS_Sales.ToString().PadLeft(3) & Space(8) & TotalMagazineSales.ToString("N0").PadLeft(5) &
                          Space(7) & BonusPay.ToString("c").PadLeft(7))
    End Sub
    Sub EndOfJob()
        Call SummaryOutput()
        Call CloseFile()
    End Sub

    Sub SummaryOutput()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine(Space(30) & "End of Magazine Report")
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine(Space(30) & "Press -ENTER- To Exit")
    End Sub

    Sub CloseFile()
        MagazineFile.Close()
        Console.ReadLine()
    End Sub
End Module
