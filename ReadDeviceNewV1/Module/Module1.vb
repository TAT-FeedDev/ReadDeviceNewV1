Imports OfficeOpenXml

Module Module1
    Sub Main()
        Dim filePath As String = "D:\MQTT\NewDevice\NEW_DEVICE_V0.xlsx"

        Dim package As New ExcelPackage(New IO.FileInfo(filePath))
        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0) ' Access the first worksheet

        Dim rows As Integer = worksheet.Dimension.End.Row
        Dim cols As Integer = worksheet.Dimension.End.Column

        For row As Integer = 1 To rows
            For col As Integer = 1 To cols
                Dim cellValue As String = worksheet.Cells(row, col).Text
                Console.WriteLine(cellValue)
            Next
        Next

        Console.ReadLine()
    End Sub
End Module
